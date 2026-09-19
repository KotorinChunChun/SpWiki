/**
 * Handing a document to a desktop application through a custom URL scheme.
 *
 * A browser cannot follow a `file:` link from an http page, so the only way out
 * is a scheme the machine has an application registered for. What travels is
 * the document's own SharePoint URL with the scheme swapped:
 *
 *   https://tenant.sharepoint.com/sites/docs/Shared%20Documents/md/README.md
 *   spmd://tenant.sharepoint.com/sites/docs/Shared%20Documents/md/README.md
 *
 * Nothing else goes with it. Where the file sits on disk depends on the machine
 * and the user, so it is the registered application's business, not this page's.
 */

/** Default scheme, used when the web part setting is empty. */
export const DEFAULT_APP_SCHEME: string = 'spmd';

/**
 * The URL to navigate to in order to hand `fileAbsoluteUrl` to the application
 * registered for `scheme`.
 *
 * The scheme is stripped of anything a URL scheme may not contain, so a setting
 * typed as `spmd://` or `spmd:` still produces a usable URL.
 */
export function buildProtocolUrl(scheme: string, fileAbsoluteUrl: string): string {
  const cleanScheme: string = scheme.replace(/[^a-zA-Z0-9+.-]/g, '') || DEFAULT_APP_SCHEME;

  if (/^https?:\/\//i.test(fileAbsoluteUrl)) {
    return fileAbsoluteUrl.replace(/^https?:\/\//i, cleanScheme + '://');
  }

  return cleanScheme + '://' + fileAbsoluteUrl.replace(/^\/+/, '');
}
