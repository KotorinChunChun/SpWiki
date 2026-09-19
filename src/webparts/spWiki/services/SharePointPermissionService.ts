import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { encodePath, joinUrl } from '../core/pathUtils';
import {
  IAssignment,
  IPermissionState,
  IPrincipal,
  IRoleDefinition,
  PrincipalKind
} from '../core/permissions';
import { escapeODataString } from './SharePointDocumentProvider';

/**
 * Reading and changing who can see a folder or a file.
 *
 * Every write here is a permission change, so each one is deliberate:
 *
 * - Breaking inheritance *copies* the current assignments
 *   (`copyRoleAssignments=true`). Nobody loses access at the moment of the
 *   break; access is then narrowed by removing entries, which is a step the
 *   reader can stop half way through without having locked anyone out.
 * - `clearSubscopes=false`, so descendants that already have permissions of
 *   their own are left exactly as they are rather than being silently
 *   flattened. Those descendants are reported instead — see
 *   {@link listBrokenDescendants}.
 */

/** Upper bound on the items scanned when looking for broken inheritance. */
const DESCENDANT_SCAN_LIMIT: number = 500;

/** Upper bound on directory results offered when granting access. */
const PICKER_RESULT_LIMIT: number = 20;

export class SharePointPermissionService {
  private readonly _client: SPHttpClient;
  private readonly _siteAbsoluteUrl: string;
  private readonly _libraryUrl: string;
  private _rolesPromise: Promise<IRoleDefinition[]> | undefined;

  public constructor(client: SPHttpClient, siteAbsoluteUrl: string, libraryServerRelativeUrl: string) {
    this._client = client;
    this._siteAbsoluteUrl = siteAbsoluteUrl.replace(/\/+$/, '');
    this._libraryUrl = libraryServerRelativeUrl.replace(/\/+$/, '');
  }

  /** Permission levels this site defines, cached for the session. */
  public getRoleDefinitions(): Promise<IRoleDefinition[]> {
    if (!this._rolesPromise) {
      this._rolesPromise = this._loadRoleDefinitions().catch((error: Error) => {
        this._rolesPromise = undefined;
        throw error;
      });
    }
    return this._rolesPromise;
  }

  /** Everything the panel shows about one folder or file. */
  public async getState(path: string, isFolder: boolean): Promise<IPermissionState> {
    const itemId: number = await this._getItemId(path, isFolder);

    const [hasUnique, assignments, brokenDescendants] = await Promise.all([
      this._hasUniqueRoleAssignments(itemId),
      this._getAssignments(itemId),
      isFolder ? this.listBrokenDescendants(path) : Promise.resolve([] as string[])
    ]);

    return {
      path: path,
      isFolder: isFolder,
      hasUnique: hasUnique,
      assignments: assignments,
      brokenDescendants: brokenDescendants
    };
  }

  /**
   * Gives the item permissions of its own, keeping the current ones.
   *
   * This is the step that stops the parent from reaching it. It is reversible
   * through {@link restoreInheritance}, but a restore throws away whatever was
   * set here, so the caller confirms first.
   */
  public async restrict(path: string, isFolder: boolean): Promise<void> {
    const itemId: number = await this._getItemId(path, isFolder);

    await this._post(
      this._itemEndpoint(itemId) +
        '/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)'
    );
  }

  /** Puts the item back under its parent's permissions. */
  public async restoreInheritance(path: string, isFolder: boolean): Promise<void> {
    const itemId: number = await this._getItemId(path, isFolder);
    await this._post(this._itemEndpoint(itemId) + '/resetroleinheritance');
  }

  /**
   * Grants a principal a permission level here.
   *
   * A directory principal that has never been used on this site has no
   * SharePoint id yet, so it is resolved first — that is what `ensureuser`
   * does, and it is why an Entra ID group can be granted access at all.
   */
  public async grant(path: string, isFolder: boolean, principal: IPrincipal, roleId: number): Promise<void> {
    const itemId: number = await this._getItemId(path, isFolder);
    const principalId: number =
      principal.id > 0 ? principal.id : await this._ensurePrincipalId(principal.loginName);

    await this._post(
      this._itemEndpoint(itemId) +
        '/roleassignments/addroleassignment(principalid=' +
        principalId +
        ',roledefid=' +
        roleId +
        ')'
    );
  }

  /**
   * Takes a principal's access away entirely.
   *
   * The whole assignment is deleted rather than each role binding removed one
   * at a time: `removeroleassignment` needs a role definition id, so revoking a
   * principal holding two levels would take two calls and could half succeed.
   */
  public async revoke(path: string, isFolder: boolean, principal: IPrincipal): Promise<void> {
    const itemId: number = await this._getItemId(path, isFolder);

    await this._post(
      this._itemEndpoint(itemId) +
        '/roleassignments/getbyprincipalid(' +
        principal.id +
        ')/deleteobject()'
    );
  }

  /**
   * Descendants of a folder that hold permissions of their own.
   *
   * These are the holes: restricting the folder does not reach them, so a
   * reader who thinks they have just made a folder private may not have. There
   * is no OData filter for `HasUniqueRoleAssignments`, so the items are read
   * and sifted here.
   */
  public async listBrokenDescendants(folderPath: string): Promise<string[]> {
    const prefix: string = joinUrl(this._libraryUrl, folderPath).toLowerCase() + '/';

    const endpoint: string =
      this._getListEndpoint() +
      '/items?$select=FileRef,HasUniqueRoleAssignments&$top=' +
      DESCENDANT_SCAN_LIMIT;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = readCollection(body);
    const broken: string[] = [];
    const libraryPrefix: string = this._libraryUrl.toLowerCase() + '/';

    for (let i: number = 0; i < rows.length; i++) {
      const fileRef: string = rows[i].FileRef || '';
      const lower: string = fileRef.toLowerCase();

      // Under the folder, but not the folder itself.
      if (lower.slice(0, prefix.length) !== prefix || lower === prefix.slice(0, -1)) {
        continue;
      }

      if (rows[i].HasUniqueRoleAssignments === true) {
        broken.push(fileRef.slice(libraryPrefix.length));
      }
    }

    broken.sort();
    return broken;
  }

  /**
   * Groups and people matching what was typed.
   *
   * The people picker endpoint is what makes Entra ID groups reachable:
   * `principalType` 15 is "all", which covers SharePoint groups, users,
   * security groups and Microsoft 365 groups in one call. The results are
   * directory entries, so most of them have no SharePoint id until they are
   * granted something — {@link grant} resolves that.
   */
  public async searchPrincipals(query: string): Promise<IPrincipal[]> {
    const endpoint: string =
      this._siteAbsoluteUrl +
      '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface' +
      '.clientPeoplePickerSearchUser';

    const parameters: string = JSON.stringify({
      queryParams: {
        __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        MaximumEntitySuggestions: PICKER_RESULT_LIMIT,
        PrincipalSource: 15,
        PrincipalType: 15,
        QueryString: query
      }
    });

    const response: SPHttpClientResponse = await this._client.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          // This endpoint predates the newer OData flavours: the body carries
          // `__metadata`, so the whole exchange has to be verbose. Asking for
          // `nometadata` back, or letting SPHttpClient send its default
          // `odata-version: 4.0`, is answered with a flat HTTP 400.
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: parameters
      }
    );

    if (!response.ok) {
      // The site's own users and groups are always readable, so a picker that
      // will not answer is a reason to offer less rather than nothing.
      return this._searchSitePrincipals(query);
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await response.json();
    // The payload is a JSON *string* inside the response, not an object.
    const raw: string = typeof body === 'string' ? body : body.value || body.d || '';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let entries: any[] = [];
    try {
      entries = JSON.parse(typeof raw === 'string' ? raw : JSON.stringify(raw));
    } catch {
      return this._searchSitePrincipals(query);
    }

    const principals: IPrincipal[] = [];
    for (let i: number = 0; i < entries.length; i++) {
      const principal: IPrincipal | undefined = toPrincipal(entries[i]);
      if (principal) {
        principals.push(principal);
      }
    }

    return principals.length > 0 ? principals : this._searchSitePrincipals(query);
  }

  /**
   * Fallback: the groups and users this site already knows about.
   *
   * Narrower than the directory — it cannot offer an Entra ID group nobody has
   * used here yet — but it needs no special permission and always answers, so
   * the panel is never a dead end.
   */
  private async _searchSitePrincipals(query: string): Promise<IPrincipal[]> {
    const wanted: string = query.toLowerCase();

    const [groups, users] = await Promise.all([
      this._getJson(this._siteAbsoluteUrl + '/_api/web/sitegroups?$select=Id,Title,LoginName').catch(
        (): undefined => undefined
      ),
      this._getJson(
        this._siteAbsoluteUrl +
          '/_api/web/siteusers?$select=Id,Title,LoginName,Email,PrincipalType&$top=200'
      ).catch((): undefined => undefined)
    ]);

    const found: IPrincipal[] = [];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const groupRows: any[] = readCollection(groups);
    for (let i: number = 0; i < groupRows.length; i++) {
      if (String(groupRows[i].Title || '').toLowerCase().indexOf(wanted) >= 0) {
        found.push({
          id: groupRows[i].Id,
          loginName: groupRows[i].LoginName || '',
          title: groupRows[i].Title || '',
          kind: 'sharePointGroup',
          email: ''
        });
      }
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const userRows: any[] = readCollection(users);
    for (let i: number = 0; i < userRows.length; i++) {
      const haystack: string =
        String(userRows[i].Title || '').toLowerCase() +
        '\n' +
        String(userRows[i].Email || '').toLowerCase() +
        '\n' +
        String(userRows[i].LoginName || '').toLowerCase();

      if (haystack.indexOf(wanted) >= 0) {
        found.push({
          id: userRows[i].Id,
          loginName: userRows[i].LoginName || '',
          title: userRows[i].Title || '',
          kind: toPrincipalKind(userRows[i].PrincipalType),
          email: userRows[i].Email || ''
        });
      }
    }

    return found.slice(0, PICKER_RESULT_LIMIT);
  }

  // ------------------------------------------------------------------ private

  private async _loadRoleDefinitions(): Promise<IRoleDefinition[]> {
    const endpoint: string = this._siteAbsoluteUrl + '/_api/web/roledefinitions?$select=Id,Name';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = readCollection(body);
    const roles: IRoleDefinition[] = [];

    for (let i: number = 0; i < rows.length; i++) {
      if (typeof rows[i].Id === 'number' && rows[i].Name) {
        roles.push({ id: rows[i].Id, name: rows[i].Name });
      }
    }

    return roles;
  }

  private async _hasUniqueRoleAssignments(itemId: number): Promise<boolean> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(
      this._itemEndpoint(itemId) + '?$select=HasUniqueRoleAssignments'
    );

    return body ? body.HasUniqueRoleAssignments === true : false;
  }

  private async _getAssignments(itemId: number): Promise<IAssignment[]> {
    const endpoint: string =
      this._itemEndpoint(itemId) +
      '/roleassignments?$expand=Member,RoleDefinitionBindings' +
      '&$select=PrincipalId,Member/Title,Member/LoginName,Member/PrincipalType,Member/Email,' +
      'RoleDefinitionBindings/Id,RoleDefinitionBindings/Name';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rows: any[] = readCollection(body);
    const assignments: IAssignment[] = [];

    for (let i: number = 0; i < rows.length; i++) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const row: any = rows[i];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const member: any = row.Member || {};

      const roles: IRoleDefinition[] = [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const bindings: any[] = readCollection(row.RoleDefinitionBindings);
      for (let b: number = 0; b < bindings.length; b++) {
        roles.push({ id: bindings[b].Id, name: bindings[b].Name || '' });
      }

      assignments.push({
        principal: {
          id: typeof row.PrincipalId === 'number' ? row.PrincipalId : 0,
          loginName: member.LoginName || '',
          title: member.Title || '',
          kind: toPrincipalKind(member.PrincipalType),
          email: member.Email || ''
        },
        roles: roles
      });
    }

    return assignments;
  }

  /** SharePoint id for a directory principal, creating the site entry if needed. */
  private async _ensurePrincipalId(loginName: string): Promise<number> {
    const response: SPHttpClientResponse = await this._client.post(
      this._siteAbsoluteUrl + '/_api/web/ensureuser',
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify({ logonName: loginName })
      }
    );

    if (!response.ok) {
      throw new Error(await describeResponse(response));
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await response.json();
    if (!body || typeof body.Id !== 'number') {
      throw new Error('ensureuser returned no id for ' + loginName);
    }

    return body.Id;
  }

  /** List item id of a folder or a file; both are securable objects. */
  private async _getItemId(path: string, isFolder: boolean): Promise<number> {
    const serverRelativeUrl: string = joinUrl(this._libraryUrl, path);
    const accessor: string = isFolder
      ? "/_api/web/GetFolderByServerRelativeUrl('"
      : "/_api/web/GetFileByServerRelativeUrl('";

    const endpoint: string =
      this._siteAbsoluteUrl +
      accessor +
      escapeODataString(encodePath(serverRelativeUrl)) +
      "')/ListItemAllFields?$select=Id";

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: any = await this._getJson(endpoint);
    if (!body || typeof body.Id !== 'number') {
      throw new Error('No list item for ' + path);
    }

    return body.Id;
  }

  private _itemEndpoint(itemId: number): string {
    return this._getListEndpoint() + '/items(' + itemId + ')';
  }

  private _getListEndpoint(): string {
    return (
      this._siteAbsoluteUrl +
      "/_api/web/getList('" +
      escapeODataString(encodePath(this._libraryUrl)) +
      "')"
    );
  }

  /**
   * A method call with no body.
   *
   * Deliberately without a `Content-Type`: these endpoints take their arguments
   * in the URL and the body is empty, and declaring an empty body to be JSON is
   * what the working calls elsewhere in this project avoid doing.
   */
  private async _post(endpoint: string): Promise<void> {
    const options: ISPHttpClientOptions = {
      headers: { Accept: 'application/json;odata=nometadata' },
      body: ''
    };

    const response: SPHttpClientResponse = await this._client.post(
      endpoint,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(await describeResponse(response));
    }
  }

  /**
   * A read.
   *
   * `no-cache` because these are read again immediately after a change: a
   * cached `HasUniqueRoleAssignments` would show the state from before the
   * break and make a change that worked look like one that did nothing.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async _getJson(endpoint: string): Promise<any> {
    const response: SPHttpClientResponse = await this._client.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Cache-Control': 'no-cache'
        }
      }
    );

    if (!response.ok) {
      throw new Error(await describeResponse(response));
    }

    return response.json();
  }
}

/**
 * SharePoint's `PrincipalType` flags, reduced to what the panel distinguishes.
 * 1 user, 2 distribution list, 4 security group, 8 SharePoint group.
 */
function toPrincipalKind(principalType: number): PrincipalKind {
  if (principalType === 8) {
    return 'sharePointGroup';
  }
  return principalType === 1 ? 'user' : 'securityGroup';
}

/** One people picker entry, or undefined when it is not something grantable. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function toPrincipal(entry: any): IPrincipal | undefined {
  if (!entry || !entry.Key) {
    return undefined;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const props: any = entry.EntityData || {};
  const type: string = String(entry.EntityType || '');

  let kind: PrincipalKind = 'user';
  if (type === 'SPGroup') {
    kind = 'sharePointGroup';
  } else if (type === 'FormsRole' || type === 'SecGroup' || props.PrincipalType === 'SecurityGroup') {
    kind = 'securityGroup';
  }

  return {
    // A directory entry has no site id yet; `grant` resolves it when used.
    id: kind === 'sharePointGroup' && props.SPGroupID ? Number(props.SPGroupID) : 0,
    loginName: String(entry.Key),
    title: String(entry.DisplayText || entry.Key),
    kind: kind,
    email: String(props.Email || '')
  };
}

/** The status line plus whatever SharePoint said in the body. */
async function describeResponse(response: SPHttpClientResponse): Promise<string> {
  const status: string = 'HTTP ' + response.status + ' ' + response.statusText;

  let body: string;
  try {
    body = await response.text();
  } catch {
    return status;
  }

  if (body === '') {
    return status;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let parsed: any;
  try {
    parsed = JSON.parse(body);
  } catch {
    return status + ': ' + (body.length > 300 ? body.slice(0, 300) : body);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const error: any = parsed ? parsed['odata.error'] || parsed.error : undefined;
  if (error && error.message && typeof error.message.value === 'string') {
    return status + ': ' + error.message.value;
  }

  return status;
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
