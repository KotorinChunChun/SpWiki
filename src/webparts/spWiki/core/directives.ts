/**
 * Special headings that expand into a generated list.
 *
 * `## 目次` established the shape: the writer types a heading, the viewer puts
 * something underneath it. These follow the same rule and take an argument
 * after a colon, so one document can carry several of them:
 *
 *   ## タグページ一覧:設計       every page tagged 設計
 *   ## サブページ一覧:guide      every page under a folder
 *   ## ページツリー              where this page sits, and its neighbours
 *   ## 関連ページ一覧:install    every page matching a keyword
 *
 * Finding the markers is separate from filling them in: the marker is in the
 * document, but the library listing that answers it only the host can fetch.
 * Everything here is pure DOM work over data handed in, so it can be tested
 * without SharePoint.
 */

import { ITreeFolder, buildTree, isEmptyTree, renderTree } from './pageTree';
import { dirname, normalizePath } from './pathUtils';
import { IRelatedPage } from './relatedPages';

const HEADING_SELECTOR: string = 'h1, h2, h3, h4, h5, h6';

export type DirectiveKind = 'tagPages' | 'subPages' | 'pageTree' | 'relatedPages';

/** Class of a generated block, so it can be styled and recognised again. */
export const DIRECTIVE_CLASS: string = 'mdv-directive';

/** Marks the entry of the page being read inside a generated list. */
export const DIRECTIVE_CURRENT_CLASS: string = 'mdv-directive-current';

interface IMarker {
  kind: DirectiveKind;
  text: string;
  /** False for a marker that stands alone, like ページツリー. */
  takesArgument: boolean;
}

/**
 * Longest first: 「タグページ一覧」 and 「サブページ一覧」 both end in
 * 「ページ一覧」, so a shorter marker must never be tested first.
 */
const MARKERS: IMarker[] = [
  { kind: 'tagPages', text: 'タグページ一覧', takesArgument: true },
  { kind: 'subPages', text: 'サブページ一覧', takesArgument: true },
  { kind: 'relatedPages', text: '関連ページ一覧', takesArgument: true },
  { kind: 'pageTree', text: 'ページツリー', takesArgument: false }
];

export interface IDirective {
  kind: DirectiveKind;
  /** What followed the colon, trimmed. Empty when the marker takes none. */
  argument: string;
  /** Heading the generated block is inserted after. */
  heading: Element;
}

/**
 * Every special heading in the document, in document order.
 *
 * A heading that merely contains a marker is not one: `サブページ一覧について`
 * is prose about the feature. The whole heading has to be the marker, or the
 * marker followed by a colon and an argument.
 */
export function findDirectives(container: HTMLElement): IDirective[] {
  const headings: NodeListOf<Element> = container.querySelectorAll(HEADING_SELECTOR);
  const found: IDirective[] = [];

  for (let i: number = 0; i < headings.length; i++) {
    const directive: IDirective | undefined = parseDirective(headings[i]);
    if (directive) {
      found.push(directive);
    }
  }

  return found;
}

/** The directive a heading carries, or undefined when it is ordinary prose. */
export function parseDirective(heading: Element): IDirective | undefined {
  // Both the ASCII colon and the full width one people actually type.
  const text: string = trim(heading.textContent || '').replace(/：/g, ':');

  for (let i: number = 0; i < MARKERS.length; i++) {
    const marker: IMarker = MARKERS[i];

    if (text === marker.text) {
      return { kind: marker.kind, argument: '', heading: heading };
    }

    if (marker.takesArgument && text.slice(0, marker.text.length + 1) === marker.text + ':') {
      const argument: string = trim(text.slice(marker.text.length + 1));
      if (argument !== '') {
        return { kind: marker.kind, argument: argument, heading: heading };
      }
    }
  }

  return undefined;
}

export interface IDirectiveStrings {
  /** Shown in place of a list that came out empty. */
  empty: string;
  /** Name of the library root in the page tree. */
  rootLabel: string;
}

export interface IDirectiveContext {
  /** Every viewable document in the library, with its tags. */
  pages: IRelatedPage[];
  /** Library relative path of the document being read. */
  currentPath: string;
  /** Address of a document, so the entries are real links. */
  buildDocUrl(path: string): string;
  /** Attribute the host's click handling looks for on an internal link. */
  linkAttribute: string;
  strings: IDirectiveStrings;
}

/** Fills in every directive found in the container. */
export function renderDirectives(
  container: HTMLElement,
  directives: IDirective[],
  context: IDirectiveContext
): void {
  for (let i: number = 0; i < directives.length; i++) {
    renderDirective(directives[i], context);
  }
}

function renderDirective(directive: IDirective, context: IDirectiveContext): void {
  const heading: Element = directive.heading;
  if (!heading.parentNode) {
    return;
  }

  const doc: Document = heading.ownerDocument as Document;
  const block: HTMLElement = doc.createElement('div');
  block.className = DIRECTIVE_CLASS + ' ' + DIRECTIVE_CLASS + '-' + directive.kind;

  if (directive.kind === 'pageTree') {
    block.appendChild(buildPageTree(doc, context));
  } else if (directive.kind === 'subPages') {
    // A folder listing is a shape, not a list: a flat run of paths hides which
    // file belongs to which folder, which is the only reason to ask for it.
    block.appendChild(buildSubPages(doc, directive, context));
  } else {
    const matches: IRelatedPage[] = selectPages(directive, context);

    block.appendChild(
      matches.length === 0
        ? buildEmpty(doc, context.strings.empty)
        : buildPageList(doc, matches, context)
    );
  }

  // Re-running over an already expanded document must not stack two blocks up.
  const existing: Node | null = heading.nextSibling;
  if (existing && existing.nodeType === 1 && isDirectiveBlock(existing as Element)) {
    heading.parentNode.replaceChild(block, existing);
    return;
  }

  heading.parentNode.insertBefore(block, heading.nextSibling);
}

function isDirectiveBlock(element: Element): boolean {
  return (element.className || '').indexOf(DIRECTIVE_CLASS) >= 0;
}

/** The pages one list directive asks for, in path order. */
export function selectPages(directive: IDirective, context: IDirectiveContext): IRelatedPage[] {
  if (directive.kind === 'tagPages') {
    return sortByPath(filterByTag(context.pages, directive.argument));
  }

  if (directive.kind === 'subPages') {
    return sortByPath(filterByFolder(context.pages, directive.argument, context.currentPath));
  }

  if (directive.kind === 'relatedPages') {
    return sortByPath(filterByKeyword(context.pages, directive.argument));
  }

  return [];
}

/** Pages carrying the tag, compared like tags are compared everywhere else. */
function filterByTag(pages: IRelatedPage[], tag: string): IRelatedPage[] {
  const wanted: string = trim(tag).toLowerCase();
  const matches: IRelatedPage[] = [];

  for (let i: number = 0; i < pages.length; i++) {
    const tags: string[] = pages[i].tags;
    for (let t: number = 0; t < tags.length; t++) {
      if (trim(tags[t]).toLowerCase() === wanted) {
        matches.push(pages[i]);
        break;
      }
    }
  }

  return matches;
}

/**
 * Pages under a folder, named either way round.
 *
 * A leading `/` means the library root — `/guide` is the top level `guide`
 * wherever the reader currently is. Anything else is relative to the folder
 * holding the document, so a page in `guide/` can point at `api` and mean
 * `guide/api`. `.` and `..` work, which is what makes `../reference` possible.
 */
export function resolveFolder(argument: string, currentPath: string): string {
  const raw: string = trim(argument).replace(/\/+$/, '');

  if (raw === '' || raw === '/') {
    return '';
  }

  if (raw.charAt(0) === '/') {
    return normalizePath(raw.slice(1));
  }

  const base: string = dirname(currentPath);
  return normalizePath(base === '' ? raw : base + '/' + raw);
}

function filterByFolder(
  pages: IRelatedPage[],
  argument: string,
  currentPath: string
): IRelatedPage[] {
  const folder: string = resolveFolder(argument, currentPath);
  const prefix: string = folder === '' ? '' : folder.toLowerCase() + '/';
  const matches: IRelatedPage[] = [];

  for (let i: number = 0; i < pages.length; i++) {
    const path: string = pages[i].path.toLowerCase();
    if (prefix === '' || path.slice(0, prefix.length) === prefix) {
      matches.push(pages[i]);
    }
  }

  return matches;
}

/**
 * Pages matching a keyword.
 *
 * The match is over the file name, the folders leading to it and the tags —
 * everything the library listing already carries. It deliberately does not open
 * every document to look inside: that is one request per file, and it is what
 * the search page is for.
 */
function filterByKeyword(pages: IRelatedPage[], keyword: string): IRelatedPage[] {
  const wanted: string = trim(keyword).toLowerCase();
  if (wanted === '') {
    return [];
  }

  const matches: IRelatedPage[] = [];

  for (let i: number = 0; i < pages.length; i++) {
    const page: IRelatedPage = pages[i];
    let haystack: string = page.path.toLowerCase();

    for (let t: number = 0; t < page.tags.length; t++) {
      haystack = haystack + '\n' + page.tags[t].toLowerCase();
    }

    if (haystack.indexOf(wanted) >= 0) {
      matches.push(page);
    }
  }

  return matches;
}

/**
 * Where the current page sits: the folders leading down to it, and what shares
 * its folder.
 *
 * The chain above is drawn as folder rows so the reader can see the route, and
 * the documents of the current folder hang off the last of them.
 */
function buildPageTree(doc: Document, context: IDirectiveContext): HTMLElement {
  const folder: string = dirname(context.currentPath);
  const parts: string[] = folder === '' ? [] : folder.split('/');

  const wrapper: HTMLElement = doc.createElement('div');
  wrapper.className = 'mdv-tree';

  // The route from the root, one row per level, each indented under the last.
  let indent: string = '';
  wrapper.appendChild(treeRow(doc, '', context.strings.rootLabel, 'mdv-tree-label'));

  let walked: string = '';
  for (let i: number = 0; i < parts.length; i++) {
    walked = walked === '' ? parts[i] : walked + '/' + parts[i];
    indent = i === 0 ? '' : indent + '   ';
    wrapper.appendChild(
      treeFolderLink(doc, indent + '└─ ', parts[i] + '/', walked, context)
    );
  }

  const childIndent: string = parts.length === 0 ? '' : indent + '   ';
  const siblings: IRelatedPage[] = sortByPath(pagesDirectlyIn(context.pages, folder));

  if (siblings.length === 0) {
    wrapper.appendChild(treeRow(doc, childIndent + '└─ ', context.strings.empty, 'mdv-related-empty'));
    return wrapper;
  }

  for (let i: number = 0; i < siblings.length; i++) {
    const last: boolean = i === siblings.length - 1;
    wrapper.appendChild(
      treePageRow(doc, childIndent + (last ? '└─ ' : '├─ '), siblings[i], context)
    );
  }

  return wrapper;
}

/** `サブページ一覧:...` — the folder and everything under it, in shape. */
function buildSubPages(
  doc: Document,
  directive: IDirective,
  context: IDirectiveContext
): HTMLElement {
  const root: string = resolveFolder(directive.argument, context.currentPath);
  const tree: ITreeFolder = buildTree(context.pages, root);

  if (isEmptyTree(tree)) {
    return buildEmpty(doc, context.strings.empty);
  }

  return renderTree(doc, tree, {
    buildDocUrl: context.buildDocUrl,
    linkAttribute: context.linkAttribute,
    currentPath: context.currentPath,
    rootLabel: root === '' ? context.strings.rootLabel : root + '/',
    linkFolders: true
  });
}

function treeRow(doc: Document, prefix: string, label: string, className: string): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-tree-row';
  row.appendChild(treePrefix(doc, prefix));

  const text: HTMLElement = doc.createElement('span');
  text.className = className;
  text.textContent = label;
  row.appendChild(text);

  return row;
}

function treeFolderLink(
  doc: Document,
  prefix: string,
  label: string,
  path: string,
  context: IDirectiveContext
): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-tree-row mdv-tree-folder-row';
  row.appendChild(treePrefix(doc, prefix));

  const link: HTMLElement = doc.createElement('a');
  link.className = 'mdv-tree-folder-link';
  link.setAttribute('href', context.buildDocUrl(path));
  link.setAttribute(context.linkAttribute, path);
  link.textContent = label;
  row.appendChild(link);

  return row;
}

function treePageRow(
  doc: Document,
  prefix: string,
  page: IRelatedPage,
  context: IDirectiveContext
): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-tree-row';

  if (page.path.toLowerCase() === context.currentPath.toLowerCase()) {
    row.className = row.className + ' ' + DIRECTIVE_CURRENT_CLASS;
  }

  row.appendChild(treePrefix(doc, prefix));

  const link: HTMLElement = doc.createElement('a');
  link.className = 'mdv-related-link';
  link.setAttribute('href', context.buildDocUrl(page.path));
  link.setAttribute(context.linkAttribute, page.path);
  link.textContent = page.name;
  row.appendChild(link);

  return row;
}

function treePrefix(doc: Document, prefix: string): HTMLElement {
  const span: HTMLElement = doc.createElement('span');
  span.className = 'mdv-tree-prefix';
  // Decoration; a screen reader must not read the box characters out.
  span.setAttribute('aria-hidden', 'true');
  span.textContent = prefix;
  return span;
}

/** Documents sitting in the folder itself, not in a folder below it. */
function pagesDirectlyIn(pages: IRelatedPage[], folder: string): IRelatedPage[] {
  const matches: IRelatedPage[] = [];

  for (let i: number = 0; i < pages.length; i++) {
    if (dirname(pages[i].path).toLowerCase() === folder.toLowerCase()) {
      matches.push(pages[i]);
    }
  }

  return matches;
}

function buildPageList(
  doc: Document,
  pages: IRelatedPage[],
  context: IDirectiveContext
): HTMLElement {
  const list: HTMLElement = doc.createElement('ul');
  list.className = 'mdv-related-list';

  for (let i: number = 0; i < pages.length; i++) {
    list.appendChild(pageItem(doc, pages[i], context, true));
  }

  return list;
}

/** One entry: a link the host navigates in place, plus the folder it lives in. */
function pageItem(
  doc: Document,
  page: IRelatedPage,
  context: IDirectiveContext,
  showFolder: boolean
): HTMLElement {
  const item: HTMLElement = doc.createElement('li');
  const isCurrent: boolean = page.path.toLowerCase() === context.currentPath.toLowerCase();

  if (isCurrent) {
    item.className = DIRECTIVE_CURRENT_CLASS;
  }

  const link: HTMLElement = doc.createElement('a');
  link.className = 'mdv-related-link';
  link.setAttribute('href', context.buildDocUrl(page.path));
  link.setAttribute(context.linkAttribute, page.path);
  link.textContent = page.name;
  item.appendChild(link);

  const folder: string = dirname(page.path);
  if (showFolder && folder !== '') {
    const label: HTMLElement = doc.createElement('span');
    label.className = 'mdv-related-folder';
    label.textContent = folder + '/';
    item.appendChild(label);
  }

  return item;
}

function buildEmpty(doc: Document, message: string): HTMLElement {
  const paragraph: HTMLElement = doc.createElement('p');
  paragraph.className = 'mdv-related-empty';
  paragraph.textContent = message;
  return paragraph;
}

function sortByPath(pages: IRelatedPage[]): IRelatedPage[] {
  return pages.slice().sort((a: IRelatedPage, b: IRelatedPage): number => {
    const x: string = a.path.toLowerCase();
    const y: string = b.path.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  });
}

/** Trims whitespace. JavaScript's \s covers U+3000 IDEOGRAPHIC SPACE too. */
function trim(value: string): string {
  return value.replace(/^\s+|\s+$/g, '');
}
