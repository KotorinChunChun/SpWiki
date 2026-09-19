/**
 * Path helpers for resolving links relative to the document that contains
 * them. All paths handled here are *library relative* — that is,
 * relative to the root folder of the document library (e.g. `guide/install.md`),
 * never server relative and never absolute.
 *
 * The code deliberately sticks to ES5 string/array APIs because the SPFx
 * TypeScript configuration targets ES5.
 */

import { DocumentFormat } from './types';

/** A link split into its path / query / hash parts. */
export interface ISplitLink {
  /** The path portion, without query string or fragment. */
  path: string;
  /** The query string including the leading `?`, or an empty string. */
  query: string;
  /** The fragment including the leading `#`, or an empty string. */
  hash: string;
}

const SCHEME_RE: RegExp = /^[a-zA-Z][a-zA-Z0-9+.-]*:/;
const MARKDOWN_RE: RegExp = /\.(md|markdown)$/i;
const HTML_RE: RegExp = /\.(html|htm)$/i;

/**
 * True when the link points outside the document library: it either carries an
 * explicit scheme (`https:`, `mailto:`, `data:`, ...) or is protocol relative
 * (`//host/path`).
 */
export function isExternalLink(link: string): boolean {
  if (!link) {
    return false;
  }
  if (link.slice(0, 2) === '//') {
    return true;
  }
  return SCHEME_RE.test(link);
}

/** True when the path looks like a Markdown document. */
export function isMarkdownPath(path: string): boolean {
  return MARKDOWN_RE.test(path);
}

/** True when the path looks like an HTML document. */
export function isHtmlPath(path: string): boolean {
  return HTML_RE.test(path);
}

/**
 * True when the viewer renders this file itself, rather than handing it to
 * SharePoint as a plain download. This is the single place that decides which
 * extensions the viewer is responsible for.
 */
export function isDocumentPath(path: string): boolean {
  return isMarkdownPath(path) || isHtmlPath(path);
}

/** How the source text of a document has to be parsed, judged by its name. */
export function documentFormat(path: string): DocumentFormat {
  return isHtmlPath(path) ? 'html' : 'markdown';
}

/** Splits `guide/install.md?x=1#section` into its three parts. */
export function splitLink(link: string): ISplitLink {
  let rest: string = link;
  let hash: string = '';
  let query: string = '';

  const hashIndex: number = rest.indexOf('#');
  if (hashIndex >= 0) {
    hash = rest.slice(hashIndex);
    rest = rest.slice(0, hashIndex);
  }

  const queryIndex: number = rest.indexOf('?');
  if (queryIndex >= 0) {
    query = rest.slice(queryIndex);
    rest = rest.slice(0, queryIndex);
  }

  return { path: rest, query: query, hash: hash };
}

/**
 * Collapses `.`, `..` and empty segments. `..` that would climb above the
 * library root is dropped, so the result can never escape the library.
 */
export function normalizePath(path: string): string {
  const segments: string[] = path.replace(/\\/g, '/').split('/');
  const out: string[] = [];

  for (let i: number = 0; i < segments.length; i++) {
    const segment: string = segments[i];
    if (segment === '' || segment === '.') {
      continue;
    }
    if (segment === '..') {
      if (out.length > 0) {
        out.pop();
      }
      continue;
    }
    out.push(segment);
  }

  return out.join('/');
}

/** The folder containing `path`, or an empty string for library-root files. */
export function dirname(path: string): string {
  const normalized: string = normalizePath(path);
  const index: number = normalized.lastIndexOf('/');
  return index < 0 ? '' : normalized.slice(0, index);
}

/**
 * Resolves a link found inside `currentPath` into a library relative path.
 *
 * `install.md` inside `guide/index.md` becomes `guide/install.md`;
 * `../README.md` inside `guide/install.md` becomes `README.md`.
 * A leading `/` means "from the library root".
 */
export function resolveRelativePath(currentPath: string, link: string): string {
  if (link.charAt(0) === '/' || link.charAt(0) === '\\') {
    return normalizePath(link);
  }

  const folder: string = dirname(currentPath);
  return normalizePath(folder === '' ? link : folder + '/' + link);
}

/** Percent-encodes every segment of a path while keeping the separators. */
export function encodePath(path: string): string {
  const segments: string[] = path.split('/');
  const encoded: string[] = [];
  for (let i: number = 0; i < segments.length; i++) {
    encoded.push(encodeURIComponent(segments[i]));
  }
  return encoded.join('/');
}

/**
 * True when the path names a folder rather than a file.
 *
 * A folder has no extension on its last segment. Anything with a dot in the
 * last segment is treated as a file, so `report.v2` is not mistaken for one and
 * `web.config` can never reach the reader through `?file=`.
 */
export function isFolderPath(path: string): boolean {
  const normalized: string = normalizePath(path);
  if (normalized === '') {
    return false;
  }

  const last: string = normalized.slice(normalized.lastIndexOf('/') + 1);
  return last !== '' && last.indexOf('.') < 0;
}

/** Joins a base URL and a library relative path, avoiding duplicate slashes. */
export function joinUrl(base: string, path: string): string {
  const left: string = base.replace(/\/+$/, '');
  const right: string = path.replace(/^\/+/, '');
  return right === '' ? left : left + '/' + right;
}

/**
 * Normalizes a user supplied document path (typically the `?file=` query value)
 * into a safe library relative path. Returns an empty string when the value is
 * unusable — an absolute URL, or a file the viewer does not render, so that a
 * configuration file cannot be pulled up as text through the query string.
 *
 * `allowFolder` additionally accepts a path with no extension at all, which is
 * how a folder is opened as a page. It stays opt-in because the check is what
 * stops `?file=web.config` — a name without an extension is not a file this
 * viewer would ever read, only a folder it would list.
 */
export function sanitizeDocPath(raw: string, allowFolder?: boolean): string {
  if (!raw) {
    return '';
  }

  let decoded: string = raw;
  try {
    decoded = decodeURIComponent(raw);
  } catch {
    // Malformed percent-encoding: fall back to the raw value.
  }

  if (isExternalLink(decoded)) {
    return '';
  }

  const path: string = normalizePath(splitLink(decoded).path);

  if (isDocumentPath(path)) {
    return path;
  }

  return allowFolder === true && path !== '' && isFolderPath(path) ? path : '';
}
