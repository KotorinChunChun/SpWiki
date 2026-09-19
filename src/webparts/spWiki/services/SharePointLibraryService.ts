import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { encodePath, isDocumentPath, joinUrl } from '../core/pathUtils';
import { buildLibraryViewUrl } from '../core/urlUtils';
import { IRelatedPage } from '../core/relatedPages';
import {
  ITagFieldInfo,
  ITagGroup,
  MultiChoiceFormat,
  TagFieldKind,
  flattenTagGroups,
  formatTagValue,
  hasMultiChoice,
  isSystemFieldName,
  normalizeTags,
  parseTagValue,
  toEntityPropertyName,
  toTagFieldKind
} from '../core/tags';
import { INavEntry } from '../core/navPanel';
import { escapeODataString } from './SharePointDocumentProvider';

/** Item level information about the document on screen. */
export interface IDocumentMetadata {
  /** List item id, needed to write the tags back. */
  itemId: number;
  /** URL that opens the file in SharePoint's own viewer/editor. */
  editUrl: string;
  /** Absolute URL of the file itself. */
  fileUrl: string;
  /** One entry per configured tag column, in the order they were configured. */
  tagGroups: ITagGroup[];
  /** Every tag of every column, de-duplicated. */
  tags: string[];
  /** True when the item carries permissions of its own. */
  hasUniquePermissions: boolean;
}

/** The contents of one folder, as the navigation popup needs them. */
export interface IFolderListing {
  folders: INavEntry[];
  /** Only the documents the viewer can render. */
  files: INavEntry[];
}

/** Upper bound on the items fetched when building the related pages list. */
const ITEM_FETCH_LIMIT: number = 500;

/**
 * Columns that exist on every document library and are never meaningful as a
 * tag column, so they are kept out of the property pane dropdown.
 */
const FIELD_DENYLIST: string[] = [
  'ContentType',
  'FileLeafRef',
  'FileDirRef',
  'FileRef',
  'LinkFilename',
  'LinkFilenameNoMenu',
  'LinkFilename2',
  'DocIcon',
  'ItemChildCount',
  'FolderChildCount',
  'SyncClientId',
  'AppAuthor',
  'AppEditor',
  'CheckoutUser',
  'ParentVersionString',
  'ParentLeafName',
  'SharedWithDetails',
  'SMTotalSize',
  'MediaServiceOCR',
  'MediaServiceAutoTags',
  'MediaServiceGenerationTime',
  'MediaServiceEventHashCode',
  'MediaServiceLocation',
  'MediaServiceKeyPoints'
];

/**
 * List level operations against the document library: reading the columns that
 * can hold tags, reading and writing the tags of one document, and listing the
 * documents used for the related pages panel.
 *
 * Everything goes through `/_api/web/getList('<server relative url>')`, which
 * works whether the library was configured by title or by folder path.
 */
export class SharePointLibraryService {
  private readonly _client: SPHttpClient;
  private readonly _siteAbsoluteUrl: string;
  private readonly _libraryUrl: string;
  private _fieldsPromise: Promise<ITagFieldInfo[]> | undefined;
  private _pagesPromise: Promise<IRelatedPage[]> | undefined;
  private _pagesField: string = '';
  private _viewUrlPromise: Promise<string> | undefined;
  /** Filled by the same fetch that builds the page list. */
  private _restrictedFolders: string[] = [];

  public constructor(client: SPHttpClient, siteAbsoluteUrl: string, libraryServerRelativeUrl: string) {
    this._client = client;
    this._siteAbsoluteUrl = siteAbsoluteUrl.replace(/\/+$/, '');
    this._libraryUrl = libraryServerRelativeUrl.replace(/\/+$/, '');
  }

  /** Columns of the library that can sensibly hold tags. */
  public getTagFields(): Promise<ITagFieldInfo[]> {
    if (!this._fieldsPromise) {
      this._fieldsPromise = this._loadTagFields().catch((error: Error) => {
        this._fieldsPromise = undefined;
        throw error;
      });
    }
    return this._fieldsPromise;
  }

  /** Item metadata for one document, or undefined when it has no list item. */
  public async getDocumentMetadata(path: string, fields: ITagFieldInfo[]): Promise<IDocumentMetadata | undefined> {
    const serverRelativeUrl: string = joinUrl(this._libraryUrl, path);
    const endpoint: string =
      this._siteAbsoluteUrl +
      "/_api/web/GetFileByServerRelativeUrl('" +
      escapeODataString(encodePath(serverRelativeUrl)) +
      "')?$expand=ListItemAllFields";

    // Started before the item is read and cached afterwards, so the extra round
    // trip only happens on the first document of a library.
    const viewUrlPromise: Promise<string> = this._getDefaultViewUrl(serverRelativeUrl);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    if (!body) {
      return undefined;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const item: any = body.ListItemAllFields;
    if (!item || typeof item.Id !== 'number') {
      return undefined;
    }

    const groups: ITagGroup[] = [];
    for (let i: number = 0; i < fields.length; i++) {
      groups.push({ field: fields[i], tags: parseTagValue(item[fields[i].entityPropertyName]) });
    }

    return {
      itemId: item.Id,
      editUrl: this._buildEditUrl(await viewUrlPromise, serverRelativeUrl, body.UniqueId),
      fileUrl: this._getOrigin() + encodePath(serverRelativeUrl),
      tagGroups: groups,
      tags: flattenTagGroups(groups),
      hasUniquePermissions: await this._readHasUniquePermissions(item)
    };
  }

  /**
   * Item metadata for a folder opened as a page.
   *
   * A folder is a list item too, so it can carry the same tag columns as a
   * document — but it hangs off the folder endpoint rather than the file one,
   * and it has no file to edit or hand to a desktop application.
   */
  public async getFolderMetadata(
    path: string,
    fields: ITagFieldInfo[]
  ): Promise<IDocumentMetadata | undefined> {
    const serverRelativeUrl: string = joinUrl(this._libraryUrl, path);
    const endpoint: string =
      this._siteAbsoluteUrl +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      escapeODataString(encodePath(serverRelativeUrl)) +
      "')/ListItemAllFields";

    const viewUrlPromise: Promise<string> = this._getDefaultViewUrl(serverRelativeUrl);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const item: any = await this._getJson(endpoint);
    if (!item || typeof item.Id !== 'number') {
      return undefined;
    }

    const groups: ITagGroup[] = [];
    for (let i: number = 0; i < fields.length; i++) {
      groups.push({ field: fields[i], tags: parseTagValue(item[fields[i].entityPropertyName]) });
    }

    // "Edit" opens the folder in the library view, which is the only thing
    // there is to open for a folder.
    const viewUrl: string = buildLibraryViewUrl(
      this._getOrigin(),
      await viewUrlPromise,
      serverRelativeUrl
    );

    return {
      itemId: item.Id,
      editUrl: viewUrl !== '' ? viewUrl : this._getOrigin() + encodePath(serverRelativeUrl),
      fileUrl: this._getOrigin() + encodePath(serverRelativeUrl),
      tagGroups: groups,
      tags: flattenTagGroups(groups),
      hasUniquePermissions: await this._readHasUniquePermissions(item)
    };
  }

  /**
   * Whether the item carries permissions of its own.
   *
   * `HasUniqueRoleAssignments` is a *property* of the list item, not one of its
   * fields, so it is absent from the default projection of an
   * `$expand=ListItemAllFields` — the value simply comes back undefined and
   * every page looked as though it inherited. Asking the item for it directly
   * does work, which is what the second request here is for.
   *
   * The inline value is still preferred: a tenant that does return it costs
   * nothing, and the extra round trip only happens where it is really missing.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async _readHasUniquePermissions(item: any): Promise<boolean> {
    if (typeof item.HasUniqueRoleAssignments === 'boolean') {
      return item.HasUniqueRoleAssignments;
    }

    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: any = await this._getJson(
        this._getListEndpoint() + '/items(' + item.Id + ')?$select=HasUniqueRoleAssignments'
      );

      return body ? body.HasUniqueRoleAssignments === true : false;
    } catch {
      // The padlock is not worth failing a page load over.
      return false;
    }
  }

  /**
   * Sub-folders and documents of one folder, for the navigation popup.
   *
   * `$expand=Folders,Files` returns both in one request. SharePoint's own
   * `Forms` folder is dropped: it holds the list's view pages, not content.
   */
  public async listFolder(folderPath: string): Promise<IFolderListing> {
    const serverRelativeUrl: string = joinUrl(this._libraryUrl, folderPath);
    const endpoint: string =
      this._siteAbsoluteUrl +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      escapeODataString(encodePath(serverRelativeUrl)) +
      "')?$select=Folders/Name,Files/Name&$expand=Folders,Files";

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    const folders: INavEntry[] = [];
    const files: INavEntry[] = [];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const folderRows: any[] = readCollection(body ? body.Folders : undefined);
    for (let i: number = 0; i < folderRows.length; i++) {
      const name: string = folderRows[i].Name || '';
      if (name !== '' && name.toLowerCase() !== 'forms') {
        folders.push({ name: name, path: joinPath(folderPath, name) });
      }
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fileRows: any[] = readCollection(body ? body.Files : undefined);
    for (let i: number = 0; i < fileRows.length; i++) {
      const name: string = fileRows[i].Name || '';
      if (name !== '' && isDocumentPath(name)) {
        files.push({ name: name, path: joinPath(folderPath, name) });
      }
    }

    folders.sort(byName);
    files.sort(byName);

    return { folders: folders, files: files };
  }

  /**
   * Every folder in the library, library relative, in path order.
   *
   * One request for all depths: asking the folder endpoint would only give one
   * level at a time, but the list's own items include the folders, and
   * `FSObjType eq 1` is exactly them.
   */
  public async listAllFolders(): Promise<string[]> {
    const endpoint: string =
      this._getListEndpoint() +
      '/items?$select=FileRef&$filter=' +
      encodeURIComponent('FSObjType eq 1') +
      '&$top=' +
      ITEM_FETCH_LIMIT;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = readCollection(body);
    const prefix: string = this._libraryUrl.toLowerCase() + '/';
    const folders: string[] = [];

    for (let i: number = 0; i < rows.length; i++) {
      const fileRef: string = rows[i].FileRef || '';
      if (fileRef.slice(0, prefix.length).toLowerCase() !== prefix) {
        continue;
      }

      const path: string = fileRef.slice(prefix.length);
      // The list's own Forms folder holds view pages, not content.
      if (path !== '' && path.toLowerCase() !== 'forms' && indexOfString(folders, path) < 0) {
        folders.push(path);
      }
    }

    folders.sort();
    return folders;
  }

  /**
   * Creates a document with the given content and returns its library relative
   * path.
   *
   * `overwrite=false` makes SharePoint refuse rather than silently replace a
   * file that is already there, which is the safer answer for a "new page"
   * button.
   */
  public async createFile(folderPath: string, fileName: string, content: string): Promise<string> {
    const folderUrl: string = joinUrl(this._libraryUrl, folderPath);
    const endpoint: string =
      this._siteAbsoluteUrl +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      escapeODataString(encodePath(folderUrl)) +
      "')/Files/add(url='" +
      escapeODataString(encodeURIComponent(fileName)) +
      "',overwrite=false)";

    const response: SPHttpClientResponse = await this._client.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          // The body is the file itself, not OData. Saying so keeps the
          // Japanese text in the template from being mangled.
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: content
      }
    );

    if (!response.ok) {
      throw new Error('HTTP ' + response.status + ' ' + response.statusText);
    }

    return joinPath(folderPath, fileName);
  }

  /**
   * Renames or moves a document inside the library.
   *
   * `flags=1` is "overwrite off"; the call fails instead of destroying whatever
   * already sits at the destination.
   */
  public async moveFile(fromPath: string, toPath: string): Promise<void> {
    const from: string = joinUrl(this._libraryUrl, fromPath);
    const to: string = joinUrl(this._libraryUrl, toPath);
    const endpoint: string =
      this._siteAbsoluteUrl +
      "/_api/web/GetFileByServerRelativeUrl('" +
      escapeODataString(encodePath(from)) +
      "')/moveto(newurl='" +
      escapeODataString(encodePath(to)) +
      "',flags=1)";

    const response: SPHttpClientResponse = await this._client.post(
      endpoint,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' }, body: '' }
    );

    if (!response.ok) {
      throw new Error('HTTP ' + response.status + ' ' + response.statusText);
    }

    this.invalidatePages();
  }

  /**
   * Every document in the library with its tags, taken from all the configured
   * columns at once. Cached, because the related pages panel is rebuilt on
   * every navigation; call {@link invalidatePages} after a tag has been written.
   */
  public getPages(fields: ITagFieldInfo[]): Promise<IRelatedPage[]> {
    const key: string = entityPropertyNames(fields).join(',');

    if (!this._pagesPromise || this._pagesField !== key) {
      this._pagesField = key;
      this._pagesPromise = this._loadPages(fields).catch((error: Error) => {
        this._pagesPromise = undefined;
        throw error;
      });
    }
    return this._pagesPromise;
  }

  /**
   * Folders holding permissions of their own, from the last {@link getPages}.
   * Empty until that has run, since it is the same fetch.
   */
  public getRestrictedFolders(): string[] {
    return this._restrictedFolders.slice();
  }

  /** Forces the next {@link getPages} call to hit the server again. */
  public invalidatePages(): void {
    this._pagesPromise = undefined;
  }

  /**
   * Writes the tags of one document back to its list item — every configured
   * column in a single request, so a save cannot half succeed.
   *
   * A value outside a choice column's choices only survives when the column
   * allows custom values; the editor greys the "add a tag" box out otherwise,
   * rather than letting the save fail.
   */
  public async setTags(itemId: number, groups: ITagGroup[]): Promise<void> {
    const response: SPHttpClientResponse = await this._postTags(itemId, groups, 'results');

    if (response.ok) {
      this.invalidatePages();
      return;
    }

    // A collection written in the shape this tenant does not want comes back as
    // a flat 400 with no hint of which shape it wanted, so the other one is
    // worth one retry before giving up. Only a 400 qualifies: a 403 or a 412
    // means something else is wrong and retrying just hides it.
    if (response.status === 400 && hasMultiChoice(groups)) {
      const retry: SPHttpClientResponse = await this._postTags(itemId, groups, 'array');

      if (retry.ok) {
        this.invalidatePages();
        return;
      }

      throw new Error(await describeResponse(retry));
    }

    throw new Error(await describeResponse(response));
  }

  // ------------------------------------------------------------------ private

  /** One attempt at the tag write, with the collections in the given shape. */
  private _postTags(
    itemId: number,
    groups: ITagGroup[],
    multiChoiceFormat: MultiChoiceFormat
  ): Promise<SPHttpClientResponse> {
    const endpoint: string = this._getListEndpoint() + '/items(' + itemId + ')';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const payload: any = {};
    for (let i: number = 0; i < groups.length; i++) {
      const group: ITagGroup = groups[i];
      payload[group.field.entityPropertyName] = formatTagValue(
        group.tags,
        group.field.kind,
        multiChoiceFormat
      );
    }

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(payload)
    };

    return this._client.post(endpoint, SPHttpClient.configurations.v1, options);
  }

  private _getListEndpoint(): string {
    return (
      this._siteAbsoluteUrl +
      "/_api/web/getList('" +
      escapeODataString(encodePath(this._libraryUrl)) +
      "')"
    );
  }

  private async _loadTagFields(): Promise<ITagFieldInfo[]> {
    // No $select: `Choices` only exists on choice fields, and asking for it
    // across the whole collection makes SharePoint reject the query.
    const filter: string = encodeURIComponent('Hidden eq false and ReadOnlyField eq false');
    const endpoint: string = this._getListEndpoint() + '/fields?$filter=' + filter;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = readCollection(body);
    const fields: ITagFieldInfo[] = [];

    for (let i: number = 0; i < rows.length; i++) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const row: any = rows[i];
      const internalName: string = row.InternalName || '';
      const kind: TagFieldKind | undefined = toTagFieldKind(row.TypeAsString || '');

      if (!kind || internalName === '' || isSystemFieldName(internalName)) {
        continue;
      }
      if (indexOfString(FIELD_DENYLIST, internalName) >= 0) {
        continue;
      }

      fields.push({
        internalName: internalName,
        // SharePoint supplies it; the derivation is the fallback for a tenant
        // that leaves the property off the field JSON.
        entityPropertyName: row.EntityPropertyName || toEntityPropertyName(internalName),
        title: row.Title || internalName,
        kind: kind,
        choices: readStringArray(row.Choices),
        // Text columns take anything; a choice column only does when its
        // "allow custom values" setting is on.
        allowFillIn: kind === 'text' || row.FillInChoice === true
      });
    }

    fields.sort((a: ITagFieldInfo, b: ITagFieldInfo): number => {
      return a.title.toLowerCase() < b.title.toLowerCase() ? -1 : a.title.toLowerCase() > b.title.toLowerCase() ? 1 : 0;
    });

    return fields;
  }

  private async _loadPages(fields: ITagFieldInfo[]): Promise<IRelatedPage[]> {
    // The item exposes a column under its entity property name, which is not
    // the internal name for a column created with a Japanese name.
    this._restrictedFolders = [];
    const names: string[] = entityPropertyNames(fields);

    if (names.length === 0) {
      return [];
    }

    const selected: string[] = [];
    for (let i: number = 0; i < names.length; i++) {
      selected.push(encodeURIComponent(names[i]));
    }

    const endpoint: string =
      this._getListEndpoint() +
      '/items?$select=Id,FileLeafRef,FileRef,FSObjType,HasUniqueRoleAssignments,' +
      selected.join(',') +
      '&$top=' +
      ITEM_FETCH_LIMIT;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = readCollection(body);
    const pages: IRelatedPage[] = [];
    const prefix: string = this._libraryUrl.toLowerCase() + '/';

    for (let i: number = 0; i < rows.length; i++) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const row: any = rows[i];
      const fileRef: string = row.FileRef || '';

      if (fileRef.slice(0, prefix.length).toLowerCase() !== prefix) {
        continue;
      }

      let tags: string[] = [];
      for (let f: number = 0; f < names.length; f++) {
        tags = tags.concat(parseTagValue(row[names[f]]));
      }

      const path: string = fileRef.slice(prefix.length);

      // Folders are securable too, and they are the unit permissions should be
      // set on, so their lock state is collected rather than thrown away.
      if (Number(row.FSObjType) === 1) {
        if (row.HasUniqueRoleAssignments === true) {
          this._restrictedFolders.push(path);
        }
        continue;
      }

      pages.push({
        path: path,
        name: row.FileLeafRef || '',
        tags: normalizeTags(tags),
        hasUniquePermissions: row.HasUniqueRoleAssignments === true
      });
    }

    return pages;
  }

  /**
   * Server relative URL of the library's default view page, e.g.
   * `/sites/docs/Shared Documents/Forms/AllItems.aspx`. Empty when it cannot be
   * read; the caller then falls back to a less specific URL.
   *
   * It has to come from the file's *parent list*, not from the folder the web
   * part is pointed at: `Shared Documents/md` is a folder inside the
   * `Shared Documents` list, and there is no `Forms/AllItems.aspx` underneath
   * it. Going through the file works whether the setting names the list root or
   * a folder inside it.
   */
  private _getDefaultViewUrl(fileServerRelativeUrl: string): Promise<string> {
    if (!this._viewUrlPromise) {
      this._viewUrlPromise = this._resolveDefaultViewUrl(fileServerRelativeUrl).catch((): string => {
        // Let the next document try again rather than caching the failure.
        this._viewUrlPromise = undefined;
        return '';
      });
    }
    return this._viewUrlPromise;
  }

  private async _resolveDefaultViewUrl(fileServerRelativeUrl: string): Promise<string> {
    const viaFile: string | undefined = await this._readDefaultViewUrl(
      this._siteAbsoluteUrl +
        "/_api/web/GetFileByServerRelativeUrl('" +
        escapeODataString(encodePath(fileServerRelativeUrl)) +
        "')/ListItemAllFields/ParentList?$select=DefaultViewUrl"
    );

    if (viaFile) {
      return viaFile;
    }

    // The configured folder as the list root. Only works when the web part is
    // pointed at the library itself, which is the common case.
    const viaList: string | undefined = await this._readDefaultViewUrl(
      this._getListEndpoint() + '?$select=DefaultViewUrl'
    );

    return viaList || '';
  }

  private async _readDefaultViewUrl(endpoint: string): Promise<string | undefined> {
    try {
      const response: SPHttpClientResponse = await this._client.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return undefined;
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const body: any = await response.json();
      return body && body.DefaultViewUrl ? body.DefaultViewUrl : undefined;
    } catch {
      return undefined;
    }
  }

  /**
   * URL the "edit" button opens.
   *
   * The library view page with `id` and `parent` is where the document library
   * itself sends a click on a file: it opens SharePoint's built-in preview and
   * editor over the folder holding the file, which is what a `.md` or `.html`
   * file needs. `Doc.aspx?sourcedoc=` hands the file to the Office viewer
   * instead, so it is only kept as a fallback for when the parent list could
   * not be read; the plain file URL below that just downloads it.
   */
  private _buildEditUrl(defaultViewUrl: string, fileServerRelativeUrl: string, uniqueId: string): string {
    const viewUrl: string = buildLibraryViewUrl(
      this._getOrigin(),
      defaultViewUrl,
      fileServerRelativeUrl
    );

    if (viewUrl !== '') {
      return viewUrl;
    }

    if (uniqueId) {
      return (
        this._siteAbsoluteUrl +
        '/_layouts/15/Doc.aspx?sourcedoc=%7B' +
        encodeURIComponent(uniqueId) +
        '%7D&action=default&mobileredirect=true'
      );
    }

    return this._getOrigin() + encodePath(fileServerRelativeUrl);
  }

  private _getOrigin(): string {
    const match: RegExpExecArray | null = /^(https?:\/\/[^/]+)/i.exec(this._siteAbsoluteUrl);
    return match ? match[1] : '';
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async _getJson(endpoint: string): Promise<any> {
    const response: SPHttpClientResponse = await this._client.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error('HTTP ' + response.status + ' ' + response.statusText);
    }

    return response.json();
  }
}

/**
 * The status line plus whatever SharePoint said in the body.
 *
 * A bare "HTTP 400" is not something anyone can act on, and it is exactly what
 * a rejected tag write produces. SharePoint does explain itself — which column,
 * which value, which permission — but only in the response body, so it is read
 * here and put in front of the reader instead of being dropped.
 */
async function describeResponse(response: SPHttpClientResponse): Promise<string> {
  const status: string = 'HTTP ' + response.status + ' ' + response.statusText;

  let body: string;
  try {
    body = await response.text();
  } catch {
    return status;
  }

  const detail: string = readErrorMessage(body);
  return detail === '' ? status : status + ': ' + detail;
}

/**
 * The human readable part of a SharePoint error body. `odata=nometadata` nests
 * it under `odata.error`, the verbose flavour under `error`; a body that is
 * neither is returned as-is once it is short enough to read.
 */
function readErrorMessage(body: string): string {
  if (body === '') {
    return '';
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let parsed: any;
  try {
    parsed = JSON.parse(body);
  } catch {
    return body.length > 300 ? body.slice(0, 300) : body;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const error: any = parsed ? parsed['odata.error'] || parsed.error : undefined;
  if (!error) {
    return '';
  }

  if (error.message && typeof error.message.value === 'string') {
    return error.message.value;
  }

  return typeof error.message === 'string' ? error.message : '';
}

/** REST collections arrive as an array, `{ value: [] }` or `{ d: { results } }`. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function readCollection(body: any): any[] {
  if (!body) {
    return [];
  }
  if (Object.prototype.toString.call(body) === '[object Array]') {
    return body;
  }
  if (Object.prototype.toString.call(body.value) === '[object Array]') {
    return body.value;
  }
  if (body.d && Object.prototype.toString.call(body.d.results) === '[object Array]') {
    return body.d.results;
  }
  return [];
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function readStringArray(value: any): string[] {
  if (!value) {
    return [];
  }
  if (Object.prototype.toString.call(value) === '[object Array]') {
    return value;
  }
  if (Object.prototype.toString.call(value.results) === '[object Array]') {
    return value.results;
  }
  return [];
}

/** The item property names of the given columns, de-duplicated and in order. */
function entityPropertyNames(fields: ITagFieldInfo[]): string[] {
  const names: string[] = [];

  for (let i: number = 0; i < fields.length; i++) {
    const name: string = fields[i].entityPropertyName;
    if (name && indexOfString(names, name) < 0) {
      names.push(name);
    }
  }

  return names;
}

/** Case-insensitive, so a tag is not added twice in a different spelling. */
function indexOfString(values: string[], value: string): number {
  const wanted: string = value.toLowerCase();

  for (let i: number = 0; i < values.length; i++) {
    if (values[i].toLowerCase() === wanted) {
      return i;
    }
  }
  return -1;
}

/** Joins a folder and a name into a library relative path. */
function joinPath(folder: string, name: string): string {
  return folder === '' ? name : folder.replace(/\/+$/, '') + '/' + name;
}

function byName(a: INavEntry, b: INavEntry): number {
  return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : a.name.toLowerCase() > b.name.toLowerCase() ? 1 : 0;
}
