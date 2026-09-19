/**
 * Turning a flat list of library relative paths into a tree, and drawing it.
 *
 * The flat list is what the library gives us: `guide/api/rest.md` and nothing
 * about the folders above it. Readers need the shape back, because a page's
 * place in the folder structure is most of what tells them what it is.
 *
 * The drawing uses box characters rather than nested `ul` indentation. Both were
 * tried; the rules are on `pre`-like alignment because the folder a file belongs
 * to has to be readable at a glance, and bullets at four depths stop conveying
 * that.
 */

import { buildPadlockIcon } from './icons';
import { IRelatedPage } from './relatedPages';

/** A folder, with what sits directly inside it. */
export interface ITreeFolder {
  /** Library relative path of the folder; '' is the root. */
  path: string;
  /** Segment name shown in the tree; '' for the root. */
  name: string;
  folders: ITreeFolder[];
  files: IRelatedPage[];
}

/**
 * Builds the folder structure implied by the paths.
 *
 * `root` is the folder the tree is rooted at; paths outside it are ignored, and
 * the segments of it are not repeated in the output.
 */
export function buildTree(pages: IRelatedPage[], root: string): ITreeFolder {
  const prefix: string = root === '' ? '' : root.replace(/\/+$/, '') + '/';
  const tree: ITreeFolder = { path: root, name: lastSegment(root), folders: [], files: [] };

  for (let i: number = 0; i < pages.length; i++) {
    const path: string = pages[i].path;

    if (prefix !== '' && path.slice(0, prefix.length).toLowerCase() !== prefix.toLowerCase()) {
      continue;
    }

    const relative: string = prefix === '' ? path : path.slice(prefix.length);
    const segments: string[] = relative.split('/');

    let folder: ITreeFolder = tree;

    // Every segment but the last names a folder on the way down.
    for (let s: number = 0; s < segments.length - 1; s++) {
      folder = childFolder(folder, segments[s]);
    }

    folder.files.push(pages[i]);
  }

  sortFolder(tree);
  return tree;
}

function childFolder(parent: ITreeFolder, name: string): ITreeFolder {
  for (let i: number = 0; i < parent.folders.length; i++) {
    if (parent.folders[i].name.toLowerCase() === name.toLowerCase()) {
      return parent.folders[i];
    }
  }

  const child: ITreeFolder = {
    path: parent.path === '' ? name : parent.path + '/' + name,
    name: name,
    folders: [],
    files: []
  };

  parent.folders.push(child);
  return child;
}

/** Folders before files, each alphabetically — the order a file explorer uses. */
function sortFolder(folder: ITreeFolder): void {
  folder.folders.sort((a: ITreeFolder, b: ITreeFolder): number => compare(a.name, b.name));
  folder.files.sort((a: IRelatedPage, b: IRelatedPage): number => compare(a.name, b.name));

  for (let i: number = 0; i < folder.folders.length; i++) {
    sortFolder(folder.folders[i]);
  }
}

/** True when the tree holds no folder and no file. */
export function isEmptyTree(folder: ITreeFolder): boolean {
  return folder.folders.length === 0 && folder.files.length === 0;
}

export interface ITreeView {
  /** Address of a document, so every file is a real link. */
  buildDocUrl(path: string): string;
  /** Attribute the host's click handling looks for on an internal link. */
  linkAttribute: string;
  /** Path of the document on screen, marked in the tree. */
  currentPath: string;
  /** Name given to the root row; omitted leaves the root row out. */
  rootLabel?: string;
  /** Folders link to themselves, so a folder can be opened as a page. */
  linkFolders?: boolean;
  /**
   * Lower-cased paths of the folders holding permissions of their own.
   *
   * Files carry their own flag, but a folder row is synthesized from the paths
   * below it and has no item of its own here, so its state has to be handed in.
   */
  restrictedFolders?: string[];
  /** Tooltip of the padlock shown beside a restricted row. */
  lockedLabel?: string;
}

/** Marks the row of the page being read. */
export const TREE_CURRENT_CLASS: string = 'mdv-tree-current';

/**
 * Renders the tree as rows carrying their own box-drawing prefix.
 *
 * The prefix is text rather than nested lists so the rules line up exactly under
 * one another, which is the whole point of drawing them.
 */
export function renderTree(doc: Document, folder: ITreeFolder, view: ITreeView): HTMLElement {
  const wrapper: HTMLElement = doc.createElement('div');
  wrapper.className = 'mdv-tree';

  if (view.rootLabel !== undefined) {
    wrapper.appendChild(folderRow(doc, view.rootLabel, folder.path, '', view));
    appendChildren(doc, wrapper, folder, '', view);
  } else {
    appendChildren(doc, wrapper, folder, '', view);
  }

  return wrapper;
}

/**
 * One level of the tree.
 *
 * `prefix` is what every row at this level starts with — the vertical rules of
 * the levels already above it. A child adds `│  ` to it when more siblings
 * follow, and three spaces when it is the last one, which is what makes the
 * rules stop at the right place.
 */
function appendChildren(
  doc: Document,
  host: HTMLElement,
  folder: ITreeFolder,
  prefix: string,
  view: ITreeView
): void {
  for (let i: number = 0; i < folder.folders.length; i++) {
    // Files are listed after the folders, so a folder is only the last row of
    // this level when there are none.
    const last: boolean = i === folder.folders.length - 1 && folder.files.length === 0;
    const child: ITreeFolder = folder.folders[i];

    host.appendChild(folderRow(doc, child.name + '/', child.path, prefix + branch(last), view));
    appendChildren(doc, host, child, prefix + rail(last), view);
  }

  for (let i: number = 0; i < folder.files.length; i++) {
    const last: boolean = i === folder.files.length - 1;
    host.appendChild(fileRow(doc, folder.files[i], prefix + branch(last), view));
  }
}

function branch(last: boolean): string {
  return last ? '└─ ' : '├─ ';
}

function rail(last: boolean): string {
  return last ? '   ' : '│  ';
}

function folderRow(
  doc: Document,
  label: string,
  path: string,
  prefix: string,
  view: ITreeView
): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-tree-row mdv-tree-folder-row';
  row.appendChild(prefixSpan(doc, prefix));

  if (view.linkFolders && path !== '') {
    // A folder is a page too: opening it shows its own tree.
    const link: HTMLElement = doc.createElement('a');
    link.className = 'mdv-tree-folder-link';
    link.setAttribute('href', view.buildDocUrl(path));
    link.setAttribute(view.linkAttribute, path);
    link.textContent = label;
    row.appendChild(link);
  } else {
    const name: HTMLElement = doc.createElement('span');
    name.className = 'mdv-tree-label';
    name.textContent = label;
    row.appendChild(name);
  }

  if (isRestrictedFolder(path, view)) {
    row.appendChild(lockMark(doc, view));
  }

  return row;
}

function isRestrictedFolder(path: string, view: ITreeView): boolean {
  if (!view.restrictedFolders || path === '') {
    return false;
  }

  const wanted: string = path.toLowerCase();
  for (let i: number = 0; i < view.restrictedFolders.length; i++) {
    if (view.restrictedFolders[i].toLowerCase() === wanted) {
      return true;
    }
  }

  return false;
}

/**
 * The padlock beside a restricted row.
 *
 * After the name rather than before it, so the names still line up under one
 * another — the alignment is what the box drawing is for, and an icon in front
 * of some rows and not others would break it.
 */
function lockMark(doc: Document, view: ITreeView): HTMLElement {
  const mark: HTMLElement = doc.createElement('span');
  mark.className = 'mdv-tree-lock';
  mark.setAttribute('role', 'img');
  mark.setAttribute('aria-label', view.lockedLabel || 'restricted');
  mark.setAttribute('title', view.lockedLabel || 'restricted');
  mark.appendChild(buildPadlockIcon(doc, true, 12));
  return mark;
}

function fileRow(
  doc: Document,
  page: IRelatedPage,
  prefix: string,
  view: ITreeView
): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-tree-row';

  if (page.path.toLowerCase() === view.currentPath.toLowerCase()) {
    row.className = row.className + ' ' + TREE_CURRENT_CLASS;
  }

  row.appendChild(prefixSpan(doc, prefix));

  const link: HTMLElement = doc.createElement('a');
  link.className = 'mdv-related-link';
  link.setAttribute('href', view.buildDocUrl(page.path));
  link.setAttribute(view.linkAttribute, page.path);
  link.textContent = page.name;
  row.appendChild(link);

  if (page.hasUniquePermissions === true) {
    row.appendChild(lockMark(doc, view));
  }

  return row;
}

function prefixSpan(doc: Document, prefix: string): HTMLElement {
  const span: HTMLElement = doc.createElement('span');
  span.className = 'mdv-tree-prefix';
  // Not decoration a screen reader should read out one box character at a time.
  span.setAttribute('aria-hidden', 'true');
  span.textContent = prefix;
  return span;
}

function lastSegment(path: string): string {
  const index: number = path.lastIndexOf('/');
  return index < 0 ? path : path.slice(index + 1);
}

function compare(a: string, b: string): number {
  const x: string = a.toLowerCase();
  const y: string = b.toLowerCase();
  return x < y ? -1 : x > y ? 1 : 0;
}
