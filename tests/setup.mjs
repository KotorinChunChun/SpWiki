// Installs a jsdom DOM as the global environment and then loads the bundled
// core modules, so the tests exercise the very same code the web part runs.
import { JSDOM } from 'jsdom';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const dom = new JSDOM('<!doctype html><html><body></body></html>', {
  url: 'https://tenant.sharepoint.com/sites/docs/SitePages/docs.aspx'
});

globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.Node = dom.window.Node;
globalThis.Element = dom.window.Element;
globalThis.HTMLElement = dom.window.HTMLElement;
globalThis.DocumentFragment = dom.window.DocumentFragment;
globalThis.NodeFilter = dom.window.NodeFilter;
globalThis.DOMParser = dom.window.DOMParser;
globalThis.trustedTypes = dom.window.trustedTypes;

const bundleUrl = pathToFileURL(path.resolve('temp/core.bundle.mjs')).href;

export const core = await import(bundleUrl);
export const sanitize = core.createSanitizer(dom.window);
export { dom };

/**
 * Renders a document into a fresh detached container and returns it.
 *
 * The format follows `context.currentPath` exactly as it does in the web part,
 * so `render(html, { currentPath: 'page.html' })` exercises the HTML path.
 * Pass `format` to override it.
 */
export function render(source, context, format) {
  const container = dom.window.document.createElement('div');
  core.renderDocument({
    source,
    container,
    sanitize,
    format,
    context: {
      currentPath: 'README.md',
      buildDocUrl: (p) => `/sites/docs/SitePages/docs.aspx?file=${p}`,
      buildFileUrl: (p) => `https://tenant.sharepoint.com/sites/docs/Shared%20Documents/${p}`,
      ...context
    }
  });
  return container;
}

/** Renders an HTML document, the way an `.html` file in the library is shown. */
export function renderHtml(source, context) {
  return render(source, { currentPath: 'page.html', ...context });
}
