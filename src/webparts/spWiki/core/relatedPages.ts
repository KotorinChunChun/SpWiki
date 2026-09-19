/**
 * Listing documents by tag: the search behind a tag chip, and the page list it
 * renders.
 */

import { dirname, isDocumentPath } from './pathUtils';
import { ITagColor, tagColor } from './tagColors';
import { normalizeTags, sharesTag } from './tags';

/** One viewable document in the library, with the tags read from its item. */
export interface IRelatedPage {
  /** Library relative path, e.g. `guide/install.md`. */
  path: string;
  /** File name shown in the list. */
  name: string;
  tags: string[];
  /** True when the document carries permissions of its own. */
  hasUniquePermissions?: boolean;
}

function byPath(a: IRelatedPage, b: IRelatedPage): number {
  return a.path.toLowerCase() < b.path.toLowerCase() ? -1 : a.path.toLowerCase() > b.path.toLowerCase() ? 1 : 0;
}

/** The `ul` of documents the search renders. */
function buildPageList(
  doc: Document,
  pages: IRelatedPage[],
  linkAttribute: string,
  buildDocUrl: (path: string) => string
): HTMLElement {
  const list: HTMLElement = doc.createElement('ul');
  list.className = 'mdv-related-list';

  for (let i: number = 0; i < pages.length; i++) {
    const page: IRelatedPage = pages[i];
    const item: HTMLElement = doc.createElement('li');

    const link: HTMLElement = doc.createElement('a');
    link.className = 'mdv-related-link';
    link.setAttribute('href', buildDocUrl(page.path));
    link.setAttribute(linkAttribute, page.path);
    link.textContent = page.name;
    item.appendChild(link);

    const folder: string = dirname(page.path);
    if (folder !== '') {
      const folderLabel: HTMLElement = doc.createElement('span');
      folderLabel.className = 'mdv-related-folder';
      folderLabel.textContent = folder + '/';
      item.appendChild(folderLabel);
    }

    list.appendChild(item);
  }

  return list;
}

/**
 * Every document carrying `tag`, in path order. Matching ignores case and
 * surrounding space, like everywhere else tags are compared, and does not care
 * which column the tag came from: the tag bar links to a value, not to a column.
 */
export function filterPagesByTag(pages: IRelatedPage[], tag: string): IRelatedPage[] {
  const wanted: string[] = normalizeTags([tag]);
  if (wanted.length === 0) {
    return [];
  }

  const matches: IRelatedPage[] = [];

  for (let i: number = 0; i < pages.length; i++) {
    const page: IRelatedPage = pages[i];
    if (isDocumentPath(page.path) && sharesTag(page.tags, wanted)) {
      matches.push(page);
    }
  }

  return matches.sort(byPath);
}

export interface ITagSearchView {
  /** Element that receives the result list. Its content is replaced. */
  container: HTMLElement;
  /** The tag that was searched for. */
  tag: string;
  pages: IRelatedPage[];
  /** Heading, e.g. `タグ:`. The tag itself is appended as a chip. */
  title: string;
  emptyMessage: string;
  /** Attribute used by the click handler to navigate in place. */
  linkAttribute: string;
  buildDocUrl(path: string): string;
}

/** The whole-library search behind a tag chip. */
export function renderTagSearch(view: ITagSearchView): void {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  const section: HTMLElement = doc.createElement('section');
  section.className = 'mdv-search';

  const heading: HTMLElement = doc.createElement('h2');
  heading.className = 'mdv-search-title';
  heading.textContent = view.title + ' ';

  const chip: HTMLElement = doc.createElement('span');
  chip.className = 'mdv-tag-link mdv-tag-static';
  chip.textContent = view.tag;
  const colour: ITagColor = tagColor(view.tag);
  chip.style.backgroundColor = colour.background;
  chip.style.color = colour.foreground;
  heading.appendChild(chip);

  section.appendChild(heading);

  if (view.pages.length === 0) {
    const empty: HTMLElement = doc.createElement('p');
    empty.className = 'mdv-related-empty';
    empty.textContent = view.emptyMessage;
    section.appendChild(empty);
  } else {
    section.appendChild(buildPageList(doc, view.pages, view.linkAttribute, view.buildDocUrl));
  }

  view.container.appendChild(section);
}
