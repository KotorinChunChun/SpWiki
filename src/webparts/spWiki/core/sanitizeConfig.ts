/**
 * DOMPurify configuration shared by the web part and by the local test harness.
 *
 * The same vocabulary applies to Markdown and to `.html` files in the library:
 * everything the Markdown parser can produce, plus the structural elements a
 * hand-written HTML page uses. Everything else — `script`, `iframe`, `object`,
 * `embed`, `style`, `form` and any event handler attribute — is removed.
 *
 * `style` is deliberately absent from both lists. An HTML document in the
 * library is displayed with the viewer's own typography, exactly like a
 * Markdown one, so a page cannot restyle the SharePoint page around it or pull
 * a resource in through a CSS `url()`.
 */

export const ALLOWED_TAGS: string[] = [
  'p', 'br', 'hr',
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'ul', 'ol', 'li',
  'table', 'thead', 'tbody', 'tfoot', 'tr', 'th', 'td', 'caption', 'colgroup', 'col',
  'pre', 'code', 'kbd', 'samp', 'var',
  'blockquote',
  'a', 'img',
  'strong', 'b', 'em', 'i', 'del', 's', 'sup', 'sub',
  'span', 'div',
  'dl', 'dt', 'dd',
  'input', // GitHub flavoured task lists render a disabled checkbox

  // Only ever produced by an HTML document, never by the Markdown parser.
  'figure', 'figcaption',
  'section', 'article', 'aside', 'header', 'footer', 'main', 'nav', 'address',
  'abbr', 'cite', 'q', 'small', 'mark', 'u', 'time', 'wbr',
  'details', 'summary'
];

export const ALLOWED_ATTR: string[] = [
  'href', 'src', 'alt', 'title', 'class', 'id',
  'align', 'colspan', 'rowspan', 'start', 'lang', 'dir',
  'type', 'checked', 'disabled',
  'width', 'height', 'span', 'scope', 'datetime', 'open', 'value', 'reversed'
];

export const FORBID_TAGS: string[] = [
  'script', 'iframe', 'object', 'embed', 'style', 'form', 'base', 'link', 'meta'
];

/** Configuration object accepted by `DOMPurify.sanitize`. */
export interface ISanitizeConfig {
  ALLOWED_TAGS: string[];
  ALLOWED_ATTR: string[];
  FORBID_TAGS: string[];
  ALLOW_DATA_ATTR: boolean;
}

export const SANITIZE_CONFIG: ISanitizeConfig = {
  ALLOWED_TAGS: ALLOWED_TAGS,
  ALLOWED_ATTR: ALLOWED_ATTR,
  FORBID_TAGS: FORBID_TAGS,
  ALLOW_DATA_ATTR: false
};
