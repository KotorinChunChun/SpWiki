/**
 * Heading behaviour inside a rendered document: collapsing a section, and
 * copying a link that jumps straight to it.
 *
 * Everything here runs *after* sanitization, so the attributes and buttons the
 * click handling depends on cannot be forged by the document itself.
 */

import { headingLevel } from './toc';

const HEADING_SELECTOR: string = 'h1, h2, h3, h4, h5, h6';

/** Attribute carrying the heading id on the copy button. */
export const HEADING_ANCHOR_ATTRIBUTE: string = 'data-mdv-anchor';

/** Marks a heading as clickable-to-collapse. */
export const HEADING_CLASS: string = 'mdv-heading';

/** Set on a heading whose section is folded away. */
export const COLLAPSED_CLASS: string = 'mdv-collapsed';

/** Set on the elements hidden by a collapsed heading. */
export const HIDDEN_CLASS: string = 'mdv-hidden';

export interface IHeadingOptions {
  /** Tooltip of the copy button. */
  copyLabel: string;
  /** Tooltip of the heading itself. */
  toggleLabel: string;
}

/**
 * Gives every heading with an id a copy button and makes it foldable.
 *
 * The button carries the id rather than a finished URL: only the host knows the
 * address of the page the document is being shown on.
 */
export function enhanceHeadings(container: HTMLElement, options: IHeadingOptions): void {
  const doc: Document = container.ownerDocument;
  const headings: NodeListOf<Element> = container.querySelectorAll(HEADING_SELECTOR);

  for (let i: number = 0; i < headings.length; i++) {
    const heading: Element = headings[i];
    const id: string = heading.getAttribute('id') || '';

    if (id === '' || heading.className.indexOf(HEADING_CLASS) >= 0) {
      continue;
    }

    heading.className = heading.className ? heading.className + ' ' + HEADING_CLASS : HEADING_CLASS;
    heading.setAttribute('title', options.toggleLabel);

    const copy: HTMLElement = doc.createElement('button');
    copy.className = 'mdv-heading-copy';
    copy.setAttribute('type', 'button');
    copy.setAttribute(HEADING_ANCHOR_ATTRIBUTE, id);
    copy.setAttribute('title', options.copyLabel);
    copy.setAttribute('aria-label', options.copyLabel);
    copy.appendChild(buildCopyIcon(doc));

    heading.appendChild(copy);
  }
}

/** Two offset rounded squares — the usual "copy" glyph, drawn inline. */
function buildCopyIcon(doc: Document): Element {
  const NS: string = 'http://www.w3.org/2000/svg';
  const svg: Element = doc.createElementNS(NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '13');
  svg.setAttribute('height', '13');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('focusable', 'false');

  const back: Element = doc.createElementNS(NS, 'rect');
  back.setAttribute('x', '1.5');
  back.setAttribute('y', '1.5');
  back.setAttribute('width', '9');
  back.setAttribute('height', '9');
  back.setAttribute('rx', '2');
  back.setAttribute('fill', 'none');
  back.setAttribute('stroke', 'currentColor');
  back.setAttribute('stroke-width', '1.4');

  const front: Element = doc.createElementNS(NS, 'rect');
  front.setAttribute('x', '5.5');
  front.setAttribute('y', '5.5');
  front.setAttribute('width', '9');
  front.setAttribute('height', '9');
  front.setAttribute('rx', '2');
  front.setAttribute('fill', 'none');
  front.setAttribute('stroke', 'currentColor');
  front.setAttribute('stroke-width', '1.4');

  svg.appendChild(back);
  svg.appendChild(front);
  return svg;
}

/**
 * Walks up from the clicked node to a heading copy button and returns the id it
 * carries, or `''`. Mirrors `findDocLinkPath`, so the hosts handle both the
 * same way.
 */
// `null` here mirrors the DOM APIs this walks (`event.target`, `parentNode`).
// eslint-disable-next-line @rushstack/no-new-null
export function findHeadingAnchor(start: Node | null, root: Node): string {
  let node: Node | null = start;

  while (node && node !== root) {
    if (node.nodeType === 1) {
      const value: string = (node as Element).getAttribute(HEADING_ANCHOR_ATTRIBUTE) || '';
      if (value !== '') {
        return value;
      }
    }
    node = node.parentNode;
  }

  return '';
}

/**
 * Folds or unfolds the section of the heading that was clicked. Returns true
 * when a heading was hit, so the host knows the click is spoken for.
 *
 * Clicks on a link inside a heading are left alone: following the link is what
 * the reader meant.
 */
// eslint-disable-next-line @rushstack/no-new-null
export function toggleHeadingAt(start: Node | null, root: HTMLElement): boolean {
  let node: Node | null = start;

  while (node && node !== root) {
    if (node.nodeType === 1) {
      const element: Element = node as Element;
      const name: string = element.nodeName.toLowerCase();

      if (name === 'a' || name === 'button') {
        return false;
      }
      if (headingLevel(element) > 0) {
        setCollapsed(element, element.className.indexOf(COLLAPSED_CLASS) < 0);
        applyCollapse(root);
        return true;
      }
    }
    node = node.parentNode;
  }

  return false;
}

function setCollapsed(heading: Element, collapsed: boolean): void {
  const without: string = heading.className
    .split(/\s+/)
    .filter((name: string): boolean => name !== '' && name !== COLLAPSED_CLASS)
    .join(' ');

  heading.className = collapsed ? without + ' ' + COLLAPSED_CLASS : without;
}

/**
 * Recomputes what is hidden, for every heading at once.
 *
 * Doing a full sweep rather than hiding the siblings of the heading that was
 * just clicked is what makes nesting behave: unfolding an `h2` must not reveal
 * the body of an `h3` inside it that the reader folded separately.
 */
export function applyCollapse(container: HTMLElement): void {
  const groups: Element[] = collectHeadingParents(container);

  for (let g: number = 0; g < groups.length; g++) {
    const parent: Element = groups[g];
    let hideUnder: number = 0;

    for (let node: Node | null = parent.firstChild; node; node = node.nextSibling) {
      if (node.nodeType !== 1) {
        continue;
      }

      const element: Element = node as Element;
      const level: number = headingLevel(element);

      if (level > 0) {
        if (hideUnder > 0 && level > hideUnder) {
          setHidden(element, true);
          continue;
        }

        setHidden(element, false);
        hideUnder = element.className.indexOf(COLLAPSED_CLASS) >= 0 ? level : 0;
        continue;
      }

      setHidden(element, hideUnder > 0);
    }
  }
}

/** Every element that directly holds at least one heading. */
function collectHeadingParents(container: HTMLElement): Element[] {
  const headings: NodeListOf<Element> = container.querySelectorAll(HEADING_SELECTOR);
  const parents: Element[] = [];

  for (let i: number = 0; i < headings.length; i++) {
    const parent: Node | null = headings[i].parentNode;
    if (parent && parent.nodeType === 1 && indexOfNode(parents, parent as Element) < 0) {
      parents.push(parent as Element);
    }
  }

  return parents;
}

function setHidden(element: Element, hidden: boolean): void {
  const has: boolean = element.className.indexOf(HIDDEN_CLASS) >= 0;

  if (hidden === has) {
    return;
  }

  if (hidden) {
    element.className = element.className ? element.className + ' ' + HIDDEN_CLASS : HIDDEN_CLASS;
  } else {
    element.className = element.className
      .split(/\s+/)
      .filter((name: string): boolean => name !== '' && name !== HIDDEN_CLASS)
      .join(' ');
  }
}

function indexOfNode(nodes: Element[], node: Element): number {
  for (let i: number = 0; i < nodes.length; i++) {
    if (nodes[i] === node) {
      return i;
    }
  }
  return -1;
}
