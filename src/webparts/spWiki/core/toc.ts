/**
 * Automatic table of contents.
 *
 * A document that carries a heading reading exactly `目次` gets the list of its
 * own headings inserted underneath it. Writers keep the marker in the source,
 * the links are generated from the ids `addHeadingIds` already assigned, so the
 * two can never drift apart.
 */

const HEADING_SELECTOR: string = 'h1, h2, h3, h4, h5, h6';

/** Text a heading must carry for the table of contents to replace it. */
export const TOC_MARKER: string = '目次';

/** Class of the generated section, so it can be styled and recognised again. */
export const TOC_CLASS: string = 'mdv-toc';

/** Heading level, 1..6, of an `h1`..`h6` element. */
export function headingLevel(heading: Element): number {
  const name: string = heading.tagName.toLowerCase();
  return name.charAt(0) === 'h' ? Number(name.charAt(1)) : 0;
}

/**
 * Inserts a nested link list of every heading that follows the `目次` marker.
 *
 * Returns true when a table of contents was generated. Documents without the
 * marker are left untouched, which is why this can run over everything.
 */
export function buildTableOfContents(container: HTMLElement): boolean {
  const headings: Element[] = toArray(container.querySelectorAll(HEADING_SELECTOR));
  let markerIndex: number = -1;

  for (let i: number = 0; i < headings.length; i++) {
    if (trim(headings[i].textContent || '') === TOC_MARKER) {
      markerIndex = i;
      break;
    }
  }

  if (markerIndex < 0) {
    return false;
  }

  const marker: Element = headings[markerIndex];
  const entries: Element[] = [];

  for (let i: number = markerIndex + 1; i < headings.length; i++) {
    if (headings[i].getAttribute('id')) {
      entries.push(headings[i]);
    }
  }

  if (entries.length === 0) {
    return false;
  }

  const doc: Document = container.ownerDocument;
  const nav: HTMLElement = doc.createElement('nav');
  nav.className = TOC_CLASS;
  nav.appendChild(buildList(doc, entries));

  // After the marker, so the heading itself stays where the writer put it.
  if (marker.parentNode) {
    marker.parentNode.insertBefore(nav, marker.nextSibling);
  }

  return true;
}

/**
 * The whole document's headings as a nested link list, for the outline the
 * toolbar button shows.
 *
 * Unlike {@link buildTableOfContents} this needs no marker in the source: the
 * button is asking "what is in this page", which is a question about any page,
 * not only the ones whose author asked for a contents list.
 *
 * Returns undefined when the document has no headings to list — a caller should
 * say so rather than show an empty box.
 */
export function buildHeadingNav(container: HTMLElement): HTMLElement | undefined {
  const headings: Element[] = toArray(container.querySelectorAll(HEADING_SELECTOR));
  const entries: Element[] = [];

  for (let i: number = 0; i < headings.length; i++) {
    // Only headings that can be linked to, which is the point of the list. The
    // generated table of contents is skipped: it is not part of the document.
    if (headings[i].getAttribute('id') && !isInsideToc(headings[i])) {
      entries.push(headings[i]);
    }
  }

  if (entries.length === 0) {
    return undefined;
  }

  const doc: Document = container.ownerDocument;
  const nav: HTMLElement = doc.createElement('nav');
  nav.className = 'mdv-outline';
  nav.appendChild(buildList(doc, entries));

  return nav;
}

function isInsideToc(element: Element): boolean {
  let node: Element | null = element.parentElement;

  while (node) {
    if ((node.className || '').indexOf(TOC_CLASS) >= 0) {
      return true;
    }
    node = node.parentElement;
  }

  return false;
}

/**
 * Builds the nested `ul` for a run of headings. A heading deeper than the
 * previous one opens a sub-list; a shallower one closes as many as needed, so
 * skipped levels (`h2` straight to `h4`) still produce valid markup.
 */
function buildList(doc: Document, headings: Element[]): HTMLElement {
  const root: HTMLElement = doc.createElement('ul');
  const lists: HTMLElement[] = [root];
  const levels: number[] = [headingLevel(headings[0])];

  for (let i: number = 0; i < headings.length; i++) {
    const heading: Element = headings[i];
    const level: number = headingLevel(heading);

    while (levels.length > 1 && level < levels[levels.length - 1]) {
      lists.pop();
      levels.pop();
    }

    if (level > levels[levels.length - 1]) {
      const parent: HTMLElement = lists[lists.length - 1];
      const host: Element | null = parent.lastChild as Element | null;
      const nested: HTMLElement = doc.createElement('ul');

      // A sub-list belongs inside the item above it; with no item above (a
      // document that starts one level too deep) it hangs off the list itself.
      if (host && host.nodeName.toLowerCase() === 'li') {
        host.appendChild(nested);
      } else {
        parent.appendChild(nested);
      }

      lists.push(nested);
      levels.push(level);
    }

    const item: HTMLElement = doc.createElement('li');
    const link: HTMLElement = doc.createElement('a');
    link.setAttribute('href', '#' + (heading.getAttribute('id') || ''));
    link.textContent = headingText(heading);
    item.appendChild(link);
    lists[lists.length - 1].appendChild(item);
  }

  return root;
}

/** The heading's own text, without anything the viewer added to it. */
function headingText(heading: Element): string {
  const clone: Element = heading.cloneNode(true) as Element;
  const extras: NodeListOf<Element> = clone.querySelectorAll('button');

  for (let i: number = 0; i < extras.length; i++) {
    const extra: Element = extras[i];
    if (extra.parentNode) {
      extra.parentNode.removeChild(extra);
    }
  }

  return trim(clone.textContent || '');
}

function trim(value: string): string {
  // `\s` covers the ideographic space, so writing it out is unnecessary.
  return value.replace(/^\s+|\s+$/g, '');
}

function toArray(nodes: NodeListOf<Element>): Element[] {
  const out: Element[] = [];
  for (let i: number = 0; i < nodes.length; i++) {
    out.push(nodes[i]);
  }
  return out;
}
