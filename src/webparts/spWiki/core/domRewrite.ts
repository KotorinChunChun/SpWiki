import {
  ISplitLink,
  isDocumentPath,
  isExternalLink,
  resolveRelativePath,
  splitLink
} from './pathUtils';
import { slugify } from './slugify';
import { DOC_LINK_ATTRIBUTE, IDocLinkContext } from './types';

/**
 * Rewrites the links and images of a rendered document so that they point at
 * the right place inside SharePoint.
 *
 * Runs *after* sanitization, which means the attributes written here cannot be
 * forged by the Markdown source: whatever `data-md-path` a document tries to
 * smuggle in is stripped by DOMPurify before this code puts its own back.
 */
export function rewriteLinks(container: HTMLElement, context: IDocLinkContext): void {
  const anchors: NodeListOf<Element> = container.querySelectorAll('a[href]');

  for (let i: number = 0; i < anchors.length; i++) {
    const anchor: Element = anchors[i];
    const href: string = anchor.getAttribute('href') || '';

    if (href === '') {
      continue;
    }

    // In-page anchors are handled by the browser against the ids we add below.
    if (href.charAt(0) === '#') {
      continue;
    }

    if (isExternalLink(href)) {
      anchor.setAttribute('target', '_blank');
      anchor.setAttribute('rel', 'noopener noreferrer');
      continue;
    }

    const parts: ISplitLink = splitLink(href);
    const resolved: string = resolveRelativePath(context.currentPath, parts.path);

    if (resolved === '') {
      continue;
    }

    if (isDocumentPath(resolved)) {
      // Real href so that middle-click, "open in new tab" and "copy link"
      // behave; the click handler intercepts plain clicks for in-place navigation.
      anchor.setAttribute(DOC_LINK_ATTRIBUTE, resolved + parts.hash);
      anchor.setAttribute('href', context.buildDocUrl(resolved) + parts.hash);
    } else {
      anchor.setAttribute('href', context.buildFileUrl(resolved) + parts.query + parts.hash);
    }
  }

  rewriteImages(container, context);
}

function rewriteImages(container: HTMLElement, context: IDocLinkContext): void {
  const images: NodeListOf<Element> = container.querySelectorAll('img[src]');

  for (let i: number = 0; i < images.length; i++) {
    const image: Element = images[i];
    const src: string = image.getAttribute('src') || '';

    image.setAttribute('loading', 'lazy');

    if (src === '' || isExternalLink(src)) {
      continue;
    }

    const parts: ISplitLink = splitLink(src);
    const resolved: string = resolveRelativePath(context.currentPath, parts.path);

    if (resolved !== '') {
      image.setAttribute('src', context.buildFileUrl(resolved) + parts.query);
    }
  }
}

/**
 * Gives every heading a stable id so that `[see below](#installation)` style
 * links work, and so a table of contents can be generated later on.
 */
export function addHeadingIds(container: HTMLElement): void {
  const headings: NodeListOf<Element> = container.querySelectorAll('h1, h2, h3, h4, h5, h6');
  const used: { [slug: string]: number } = {};

  for (let i: number = 0; i < headings.length; i++) {
    const heading: Element = headings[i];
    if (heading.getAttribute('id')) {
      continue;
    }

    const base: string = slugify(heading.textContent || '');
    if (base === '') {
      continue;
    }

    const seen: number = used[base] || 0;
    used[base] = seen + 1;
    heading.setAttribute('id', seen === 0 ? base : base + '-' + seen);
  }
}

/**
 * Walks up from the clicked node looking for a link that points at another
 * viewable document, and returns its library relative path (`''` if there is
 * none). Shared by the web part and by the local test harness so both navigate
 * exactly the same way.
 */
// `null` here mirrors the DOM APIs this walks (`event.target`, `parentNode`).
// eslint-disable-next-line @rushstack/no-new-null
export function findDocLinkPath(start: Node | null, root: Node): string {
  let node: Node | null = start;

  while (node && node !== root) {
    if (node.nodeType === 1) {
      const value: string = (node as Element).getAttribute(DOC_LINK_ATTRIBUTE) || '';
      if (value !== '') {
        return value;
      }
    }
    node = node.parentNode;
  }

  return '';
}

/**
 * True for an unmodified primary-button click — the only kind we intercept, so
 * that ctrl/cmd/middle clicks keep opening a new tab as the user expects.
 */
export function isPlainLeftClick(event: MouseEvent): boolean {
  if (event.defaultPrevented || event.button !== 0) {
    return false;
  }
  return !(event.metaKey || event.ctrlKey || event.shiftKey || event.altKey);
}
