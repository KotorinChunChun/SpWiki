import { ICodeBlockOptions, enhanceCodeBlocks } from './codeBlocks';
import { addHeadingIds, rewriteLinks } from './domRewrite';
import { IHeadingOptions, enhanceHeadings } from './headings';
import { highlightCodeBlocks } from './highlight';
import { markdownToHtml } from './markdown';
import { documentFormat } from './pathUtils';
import { buildTableOfContents } from './toc';
import { DocumentFormat, IDocLinkContext, Sanitizer } from './types';

export interface IRenderRequest {
  /** Source text of the document: Markdown, or HTML markup. */
  source: string;
  /** Element that receives the rendered document. Its content is replaced. */
  container: HTMLElement;
  /** Link resolution context for the document being rendered. */
  context: IDocLinkContext;
  /** Sanitizer to run the generated HTML through. */
  sanitize: Sanitizer;
  /**
   * How to read `source`. Left out, it follows the extension of
   * `context.currentPath`, which is what every caller wants.
   */
  format?: DocumentFormat;

  /** Tooltips for the per-heading controls; left out, they are not localized. */
  headings?: IHeadingOptions;

  /** Tooltip of the code block copy button; left out, it is not localized. */
  codeBlocks?: ICodeBlockOptions;
}

const DEFAULT_HEADING_OPTIONS: IHeadingOptions = {
  copyLabel: 'Copy link to this heading',
  toggleLabel: 'Click to fold'
};

const DEFAULT_CODE_BLOCK_OPTIONS: ICodeBlockOptions = {
  copyLabel: 'Copy this code'
};

/**
 * source -> HTML -> sanitize -> rewrite links -> highlight code.
 *
 * Markdown is parsed first; an HTML document skips that step and goes straight
 * into the sanitizer, so both end up under exactly the same rules — the element
 * vocabulary the sanitizer allows, and nothing else.
 *
 * The order matters: sanitization happens on the raw HTML string before it ever
 * touches the live DOM, and the link rewriting runs afterwards so that the
 * document cannot influence the attributes we depend on.
 */
export function renderDocument(request: IRenderRequest): void {
  const format: DocumentFormat = request.format || documentFormat(request.context.currentPath);
  const html: string = format === 'html' ? request.source : markdownToHtml(request.source);

  request.container.innerHTML = request.sanitize(html);

  addHeadingIds(request.container);
  rewriteLinks(request.container, request.context);

  // The table of contents is built from the ids assigned just above, and before
  // the headings grow their controls so it does not pick those up as text.
  buildTableOfContents(request.container);
  enhanceHeadings(request.container, request.headings || DEFAULT_HEADING_OPTIONS);

  highlightCodeBlocks(request.container);

  // After highlighting, so the wrapper does not get in the way of the
  // `pre > code` selector that drives it.
  enhanceCodeBlocks(request.container, request.codeBlocks || DEFAULT_CODE_BLOCK_OPTIONS);
}
