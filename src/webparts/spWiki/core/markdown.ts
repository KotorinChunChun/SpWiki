import { Marked } from 'marked';

/**
 * A dedicated Marked instance so that options stay local to this web part and
 * do not leak into anything else running on the page.
 */
const markedInstance: Marked = new Marked({
  gfm: true,
  breaks: false,
  pedantic: false
});

/** Converts Markdown source into (still untrusted) HTML. */
export function markdownToHtml(markdown: string): string {
  return markedInstance.parse(markdown, { async: false }) as string;
}
