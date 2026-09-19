import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { encodePath, joinUrl } from '../core/pathUtils';
import { IDocumentProvider, IDocumentResult } from '../core/types';

/** Prefix of the error raised when the configured library cannot be located. */
export const LIBRARY_NOT_FOUND: string = 'Document library not found';

/**
 * Reads Markdown documents out of a SharePoint document library using the REST
 * endpoint `/_api/web/GetFileByServerRelativeUrl(...)/$value`.
 */
export class SharePointDocumentProvider implements IDocumentProvider {
  private readonly _client: SPHttpClient;
  private readonly _siteAbsoluteUrl: string;
  private readonly _libraryName: string;
  private _libraryUrlPromise: Promise<string> | undefined;

  /**
   * @param client   SPHttpClient from the web part context.
   * @param siteAbsoluteUrl Absolute URL of the site holding the library.
   * @param libraryName Library title (`Documents`) or its folder path
   *                    relative to the site (`Shared Documents`).
   */
  public constructor(client: SPHttpClient, siteAbsoluteUrl: string, libraryName: string) {
    this._client = client;
    this._siteAbsoluteUrl = siteAbsoluteUrl.replace(/\/+$/, '');
    this._libraryName = libraryName.replace(/^\/+|\/+$/g, '');
  }

  /** Absolute URL of the site the provider reads from. */
  public get siteAbsoluteUrl(): string {
    return this._siteAbsoluteUrl;
  }

  /**
   * Server relative URL of the library root folder, e.g.
   * `/sites/docs/Shared Documents`.
   *
   * The library is looked up by title first, which is what makes a setting of
   * `Documents` work even though the folder is really called
   * `Shared Documents`. If no list with that title exists, the setting is used
   * as a folder path instead.
   */
  public getLibraryServerRelativeUrl(): Promise<string> {
    if (!this._libraryUrlPromise) {
      // Drop a failed lookup so the next navigation retries instead of
      // replaying the same rejection forever.
      this._libraryUrlPromise = this._resolveLibraryUrl().catch((error: Error) => {
        this._libraryUrlPromise = undefined;
        throw error;
      });
    }
    return this._libraryUrlPromise;
  }

  /** Absolute URL of a file inside the library, given its library relative path. */
  public getFileAbsoluteUrl(libraryServerRelativeUrl: string, path: string): string {
    const origin: string = this._getOrigin();
    return origin + encodePath(joinUrl(libraryServerRelativeUrl, path));
  }

  public async getDocument(path: string): Promise<IDocumentResult> {
    let libraryUrl: string;
    try {
      libraryUrl = await this.getLibraryServerRelativeUrl();
    } catch (error) {
      return { kind: 'error', message: describeError(error) };
    }

    const serverRelativeUrl: string = joinUrl(libraryUrl, path);
    const endpoint: string =
      this._siteAbsoluteUrl +
      "/_api/web/GetFileByServerRelativeUrl('" +
      escapeODataString(encodePath(serverRelativeUrl)) +
      "')/$value";

    try {
      const response: SPHttpClientResponse = await this._client.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'text/plain' } }
      );

      if (response.status === 404) {
        return { kind: 'notfound' };
      }

      if (!response.ok) {
        return {
          kind: 'error',
          message: 'HTTP ' + response.status + ' ' + response.statusText
        };
      }

      return { kind: 'ok', text: await response.text() };
    } catch (error) {
      return { kind: 'error', message: describeError(error) };
    }
  }

  private async _resolveLibraryUrl(): Promise<string> {
    // 1. The setting as a list title. This is what makes "Documents" work even
    //    though the folder is called "Shared Documents" — and, on a Japanese
    //    site, what makes the displayed title "ドキュメント" work.
    //    A value containing a slash can only be a folder path, so the lookup is
    //    skipped rather than firing a request that is bound to 404.
    if (this._libraryName.indexOf('/') < 0) {
      const byTitle: string | undefined = await this._readServerRelativeUrl(
        this._siteAbsoluteUrl +
          "/_api/web/lists/getByTitle('" +
          escapeODataString(encodeURIComponent(this._libraryName)) +
          "')/RootFolder?$select=ServerRelativeUrl"
      );

      if (byTitle) {
        return byTitle;
      }
    }

    // 2. The setting as a folder path relative to the site, e.g.
    //    "Shared Documents" or "Documents/subfolder".
    const folderUrl: string = joinUrl(this._getSiteServerRelativeUrl(), this._libraryName);
    const byPath: string | undefined = await this._readServerRelativeUrl(
      this._siteAbsoluteUrl +
        "/_api/web/GetFolderByServerRelativeUrl('" +
        escapeODataString(encodePath(folderUrl)) +
        "')?$select=ServerRelativeUrl"
    );

    if (byPath) {
      return byPath;
    }

    // Neither worked. Saying so beats reporting every document as missing.
    throw new Error(LIBRARY_NOT_FOUND + ': ' + this._libraryName);
  }

  /** GETs an endpoint that returns a folder, and picks its ServerRelativeUrl. */
  private async _readServerRelativeUrl(endpoint: string): Promise<string | undefined> {
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
      return body && body.ServerRelativeUrl ? body.ServerRelativeUrl : undefined;
    } catch {
      return undefined;
    }
  }

  private _getOrigin(): string {
    const match: RegExpExecArray | null = /^(https?:\/\/[^/]+)/i.exec(this._siteAbsoluteUrl);
    return match ? match[1] : '';
  }

  private _getSiteServerRelativeUrl(): string {
    const origin: string = this._getOrigin();
    const path: string = this._siteAbsoluteUrl.slice(origin.length);
    return path === '' ? '/' : path;
  }
}

/** OData string literals escape a single quote by doubling it. */
export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function describeError(error: any): string {
  if (!error) {
    return 'Unknown error';
  }
  return typeof error === 'string' ? error : error.message || String(error);
}
