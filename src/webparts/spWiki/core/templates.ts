/**
 * Starting content for a page created from the viewer.
 *
 * A brand new file is opened in an editor straight away, so it is worth giving
 * the writer something to edit rather than an empty buffer: the title, and the
 * `目次` marker that turns into a table of contents on its own.
 */

import { isDocumentPath } from './pathUtils';
import { DocumentFormat } from './types';

/** The heading a new document gets: its file name without the extension. */
export function documentTitle(fileName: string): string {
  const name: string = fileName.replace(/^.*[/\\]/, '');
  const dot: number = name.lastIndexOf('.');

  return isDocumentPath(name) && dot > 0 ? name.slice(0, dot) : name;
}

/** Content of a newly created document. */
export function newDocumentTemplate(fileName: string, format: DocumentFormat): string {
  const title: string = documentTitle(fileName);

  if (format === 'html') {
    return (
      '<!doctype html>\n' +
      '<html lang="ja">\n' +
      '<head>\n' +
      '  <meta charset="utf-8">\n' +
      '  <title>' + escapeHtml(title) + '</title>\n' +
      '</head>\n' +
      '<body>\n' +
      '  <h1>' + escapeHtml(title) + '</h1>\n' +
      '\n' +
      '  <h2>目次</h2>\n' +
      '\n' +
      '  <h2>概要</h2>\n' +
      '  <p></p>\n' +
      '</body>\n' +
      '</html>\n'
    );
  }

  return '# ' + title + '\n\n## 目次\n\n## 概要\n\n';
}

/** Escapes the title for the places the HTML template drops it into. */
function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
