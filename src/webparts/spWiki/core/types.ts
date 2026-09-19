/** Attribute stamped on links that point at another viewable document. */
export const DOC_LINK_ATTRIBUTE: string = 'data-md-path';

/**
 * How the text of a document has to be treated before it is sanitized:
 * `markdown` goes through the Markdown parser, `html` is already markup.
 */
export type DocumentFormat = 'markdown' | 'html';

/**
 * Everything the renderer needs in order to turn the relative links inside a
 * document into URLs that work on a SharePoint page.
 */
export interface IDocLinkContext {
  /** Library relative path of the document being rendered, e.g. `guide/install.md`. */
  currentPath: string;

  /**
   * URL of the page that should display another document.
   * Used as the `href` so that "open in new tab" and "copy link" keep working,
   * even though normal clicks are intercepted.
   */
  buildDocUrl(path: string): string;

  /**
   * URL of a file the viewer does not render itself (images, PDFs, ...),
   * given its library relative path.
   */
  buildFileUrl(path: string): string;
}

/** Turns untrusted HTML into HTML that is safe to inject into the page. */
export type Sanitizer = (html: string) => string;

/** Outcome of a document fetch. */
export type DocumentResultKind = 'ok' | 'notfound' | 'error';

export interface IDocumentResult {
  kind: DocumentResultKind;
  /** Source text of the document, present when `kind === 'ok'`. */
  text?: string;
  /** Diagnostic message, present when `kind === 'error'`. */
  message?: string;
}

/** Reads documents out of some backing store. */
export interface IDocumentProvider {
  getDocument(path: string): Promise<IDocumentResult>;
}
