/**
 * Helpers for reading and rewriting the page query string (`?file=...`), and for
 * building the SharePoint URLs the viewer links out to.
 */

import { encodePath } from './pathUtils';

/** Reads a single parameter out of a `?a=1&b=2` style query string. */
export function getQueryParam(search: string, key: string): string {
  const query: string = search.charAt(0) === '?' ? search.slice(1) : search;
  if (query === '') {
    return '';
  }

  const pairs: string[] = query.split('&');
  for (let i: number = 0; i < pairs.length; i++) {
    const pair: string = pairs[i];
    if (pair === '') {
      continue;
    }
    const eq: number = pair.indexOf('=');
    const name: string = eq < 0 ? pair : pair.slice(0, eq);
    if (decodeQueryComponent(name) === key) {
      return eq < 0 ? '' : decodeQueryComponent(pair.slice(eq + 1));
    }
  }

  return '';
}

/**
 * Returns a new query string with `key` set to `value`, preserving every other
 * parameter and their order. Slashes are left readable so the shared URL looks
 * like `?file=guide/install.md` rather than `?file=guide%2Finstall.md`.
 */
export function setQueryParam(search: string, key: string, value: string): string {
  const query: string = search.charAt(0) === '?' ? search.slice(1) : search;
  const encodedKey: string = encodeQueryComponent(key);
  const encodedValue: string = encodeQueryComponent(value);
  const out: string[] = [];
  let replaced: boolean = false;

  if (query !== '') {
    const pairs: string[] = query.split('&');
    for (let i: number = 0; i < pairs.length; i++) {
      const pair: string = pairs[i];
      if (pair === '') {
        continue;
      }
      const eq: number = pair.indexOf('=');
      const name: string = eq < 0 ? pair : pair.slice(0, eq);
      if (decodeQueryComponent(name) === key) {
        if (!replaced) {
          out.push(encodedKey + '=' + encodedValue);
          replaced = true;
        }
        continue;
      }
      out.push(pair);
    }
  }

  if (!replaced) {
    out.push(encodedKey + '=' + encodedValue);
  }

  return '?' + out.join('&');
}

/** Returns a new query string without `key`, preserving everything else. */
export function removeQueryParam(search: string, key: string): string {
  const query: string = search.charAt(0) === '?' ? search.slice(1) : search;
  if (query === '') {
    return '';
  }

  const pairs: string[] = query.split('&');
  const out: string[] = [];

  for (let i: number = 0; i < pairs.length; i++) {
    const pair: string = pairs[i];
    if (pair === '') {
      continue;
    }
    const eq: number = pair.indexOf('=');
    const name: string = eq < 0 ? pair : pair.slice(0, eq);
    if (decodeQueryComponent(name) !== key) {
      out.push(pair);
    }
  }

  return out.length === 0 ? '' : '?' + out.join('&');
}

/**
 * URL that opens one file in SharePoint's own library view — the page a click
 * in the document library leads to, with the built-in preview/editor open over
 * the folder the file lives in.
 *
 *   origin  https://tenant.sharepoint.com
 *   view    /sites/docs/Shared Documents/Forms/AllItems.aspx
 *   file    /sites/docs/Shared Documents/md/README.md
 *   ->      https://tenant.sharepoint.com/sites/docs/Shared%20Documents/Forms/AllItems.aspx
 *           ?id=%2Fsites%2Fdocs%2FShared%20Documents%2Fmd%2FREADME.md
 *           &parent=%2Fsites%2Fdocs%2FShared%20Documents%2Fmd
 *
 * `parent` is what makes the view show the folder the file is in rather than
 * the library root, so closing the preview leaves the reader in the right
 * place. The view page belongs to the *list*, which is not necessarily the
 * folder the web part is pointed at.
 *
 * Both inputs are server relative and *decoded*; an already encoded view URL is
 * accepted too, so it can be passed straight through from REST either way.
 * Returns an empty string when there is not enough to build a URL.
 */
export function buildLibraryViewUrl(
  origin: string,
  defaultViewUrl: string,
  fileServerRelativeUrl: string
): string {
  if (!defaultViewUrl || !fileServerRelativeUrl) {
    return '';
  }

  const separator: number = fileServerRelativeUrl.lastIndexOf('/');
  const parent: string = separator <= 0 ? '/' : fileServerRelativeUrl.slice(0, separator);

  return (
    origin +
    encodePath(decodeUrlPath(defaultViewUrl)) +
    '?id=' +
    encodeURIComponent(fileServerRelativeUrl) +
    '&parent=' +
    encodeURIComponent(parent)
  );
}

/** Percent-decodes a path that may or may not already be decoded. */
function decodeUrlPath(value: string): string {
  try {
    return decodeURI(value);
  } catch {
    // Malformed percent-encoding: keep what we were given.
    return value;
  }
}

function encodeQueryComponent(value: string): string {
  // Slashes and dots are legal in a query value and keep shared URLs readable.
  return encodeURIComponent(value).replace(/%2F/g, '/');
}

function decodeQueryComponent(value: string): string {
  try {
    return decodeURIComponent(value.replace(/\+/g, ' '));
  } catch {
    return value;
  }
}
