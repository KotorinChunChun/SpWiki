/**
 * The copy button on a code block.
 *
 * Like the heading controls, this runs *after* sanitization, so the marker the
 * click handling looks for cannot be forged by the document itself.
 */

/** Marks the button, and carries nothing else: the text comes from the DOM. */
export const CODE_COPY_ATTRIBUTE: string = 'data-mdv-copy-code';

/** Wrapper added around each `pre`, so the button has something to sit in. */
export const CODE_BLOCK_CLASS: string = 'mdv-code-block';

export interface ICodeBlockOptions {
  /** Tooltip and accessible name of the button. */
  copyLabel: string;
}

/**
 * Gives every code block a copy button in its top right corner.
 *
 * The `pre` is wrapped rather than positioned against directly: a `pre` is the
 * scroll container for a long line, and a button inside a scrolling box slides
 * away with the code.
 */
export function enhanceCodeBlocks(container: HTMLElement, options: ICodeBlockOptions): void {
  const doc: Document = container.ownerDocument;
  const blocks: NodeListOf<Element> = container.querySelectorAll('pre > code');

  for (let i: number = 0; i < blocks.length; i++) {
    const pre: Element | null = blocks[i].parentElement;
    if (!pre || !pre.parentElement) {
      continue;
    }

    // Already wrapped, e.g. because the document was re-rendered in place.
    const parent: Element = pre.parentElement;
    if (parent.className.indexOf(CODE_BLOCK_CLASS) >= 0) {
      continue;
    }

    const wrapper: HTMLElement = doc.createElement('div');
    wrapper.className = CODE_BLOCK_CLASS;
    parent.insertBefore(wrapper, pre);
    wrapper.appendChild(pre);

    const button: HTMLElement = doc.createElement('button');
    button.className = 'mdv-code-copy';
    button.setAttribute('type', 'button');
    button.setAttribute(CODE_COPY_ATTRIBUTE, '');
    button.setAttribute('title', options.copyLabel);
    button.setAttribute('aria-label', options.copyLabel);
    button.appendChild(buildCopyIcon(doc));

    wrapper.appendChild(button);
  }
}

/**
 * The code of the block whose copy button was clicked, or '' when the click was
 * somewhere else.
 */
// `null` here mirrors the DOM APIs this walks (`event.target`, `parentNode`).
// eslint-disable-next-line @rushstack/no-new-null
export function findCodeToCopy(start: Node | null, root: Node): string {
  let node: Node | null = start;

  while (node && node !== root) {
    if (node.nodeType === 1) {
      const element: Element = node as Element;

      if (element.hasAttribute(CODE_COPY_ATTRIBUTE)) {
        const wrapper: Element | null = element.parentElement;
        const code: Element | null = wrapper ? wrapper.querySelector('pre > code') : null;
        return code ? code.textContent || '' : '';
      }
    }
    node = node.parentNode;
  }

  return '';
}

/** Two offset rectangles — the usual "copy" glyph, drawn rather than imported. */
function buildCopyIcon(doc: Document): SVGElement {
  const ns: string = 'http://www.w3.org/2000/svg';
  const svg: SVGElement = doc.createElementNS(ns, 'svg') as SVGElement;
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '14');
  svg.setAttribute('height', '14');
  svg.setAttribute('aria-hidden', 'true');

  const path: Element = doc.createElementNS(ns, 'path');
  path.setAttribute('fill', 'currentColor');
  path.setAttribute(
    'd',
    'M10 1H2v10h2V3h6zm1 2H5v12h10V7zm2.6 4-.6-.6V6h-.4L11 4.4V7h2.6zM6 14V4h4v4h4v6z'
  );
  svg.appendChild(path);

  return svg;
}
