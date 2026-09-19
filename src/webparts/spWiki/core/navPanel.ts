/**
 * The popup behind the hamburger button, and the rename panel behind the file
 * name.
 *
 * Both are pure builders like the rest of `core/`: they take the data and the
 * callbacks, and never touch SharePoint themselves.
 */

import { isDocumentPath, normalizePath } from './pathUtils';
import { DocumentFormat } from './types';

/** A folder or a document listed in the popup. */
export interface INavEntry {
  /** Library relative path. */
  path: string;
  /** Name shown in the list. */
  name: string;
}

/** Where the writer is taken once the page exists. */
export type NewPageEditor = 'sharepoint' | 'app';

/** What the "new page" form asks for. */
export interface INewPageRequest {
  name: string;
  format: DocumentFormat;
  editor: NewPageEditor;
}

export interface INavPanelStrings {
  /** Label of the button that opens the "new page" form. */
  newPage: string;
  /** Placeholder of the file name box. */
  newPageName: string;
  /** Label above the format choice. */
  format: string;
  /** Label above the editor choice. */
  openWith: string;
  /** The two editor choices. */
  openInSharePoint: string;
  openInApp: string;
  /** Label of the create button. */
  create: string;
  cancel: string;
  /** Heading above the folder tree. */
  location: string;
  /** Name shown for the library root in the tree. */
  root: string;
  /** Shown when the current folder holds no documents. */
  empty: string;
  close: string;
  /** Opens the search page. */
  search: string;
  /** Opens the help panel. */
  help: string;
  /**
   * Switches between the SharePoint chrome and the full screen view. The label
   * names the state it switches *to*, so it reads as an action.
   */
  chromeToggle: string;
}

export interface INavPanelView {
  container: HTMLElement;
  /**
   * Folders from the library root down to the folder of the current document,
   * outermost first. Empty when the document sits in the root.
   */
  ancestors: INavEntry[];
  /** Sub-folders of the current folder. */
  folders: INavEntry[];
  /** Documents in the current folder. */
  files: INavEntry[];
  /** Library relative path of the document on screen, highlighted in the list. */
  currentPath: string;
  strings: INavPanelStrings;
  /** Real href for a document, so it can be opened in a new tab. */
  buildDocUrl(path: string): string;
  /** A document was picked. */
  onOpen(path: string): void;
  /** A folder was picked; the host reloads the panel for it. */
  onFolder(path: string): void;
  /** Create a document from the filled in form, in the current folder. */
  onCreate(request: INewPageRequest): void;
  onClose(): void;
  /** Opens the search page; left out leaves the entry off. */
  onSearch?(): void;
  /** Opens the help panel; left out leaves the entry off. */
  onHelp?(): void;
  /** Switches the page chrome; left out leaves the entry off. */
  onChromeToggle?(): void;
}

/** Builds the popup: the new page form on top, then where you are and what is next to you. */
export function renderNavPanel(view: INavPanelView): void {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  const panel: HTMLElement = doc.createElement('div');
  panel.className = 'mdv-nav-panel';

  panel.appendChild(buildNewPage(doc, view));
  panel.appendChild(buildTree(doc, view));

  // Whole-viewer actions, below the tree: they are about the viewer rather than
  // about where you are in it, so they must not compete with the file list.
  const tools: HTMLElement = doc.createElement('div');
  tools.className = 'mdv-nav-tools';

  if (view.onSearch) {
    const onSearch: () => void = view.onSearch;
    tools.appendChild(buildToolLink(doc, view.strings.search, 'mdv-nav-search', onSearch));
  }

  if (view.onChromeToggle) {
    const onChromeToggle: () => void = view.onChromeToggle;
    tools.appendChild(
      buildToolLink(doc, view.strings.chromeToggle, 'mdv-nav-chrome', onChromeToggle)
    );
  }

  if (view.onHelp) {
    const onHelp: () => void = view.onHelp;
    tools.appendChild(buildToolLink(doc, view.strings.help, 'mdv-nav-help', onHelp));
  }

  if (tools.childNodes.length > 0) {
    panel.appendChild(tools);
  }

  const footer: HTMLElement = doc.createElement('div');
  footer.className = 'mdv-nav-footer';

  const close: HTMLElement = doc.createElement('button');
  close.className = 'mdv-btn';
  close.setAttribute('type', 'button');
  close.textContent = view.strings.close;
  close.onclick = (): void => view.onClose();
  footer.appendChild(close);

  panel.appendChild(footer);
  view.container.appendChild(panel);
}

function buildToolLink(
  doc: Document,
  label: string,
  className: string,
  onClick: () => void
): HTMLElement {
  const button: HTMLElement = doc.createElement('button');
  button.className = 'mdv-nav-tool ' + className;
  button.setAttribute('type', 'button');
  button.textContent = label;
  button.onclick = (): void => onClick();
  return button;
}

/**
 * The "new page" block: a single button, which opens the form asking for the
 * name, the format and where to edit. Asking only once the writer has said they
 * want a new page keeps the popup readable when they only came to navigate.
 */
function buildNewPage(doc: Document, view: INavPanelView): HTMLElement {
  const block: HTMLElement = doc.createElement('div');
  block.className = 'mdv-nav-new';

  const open: HTMLElement = doc.createElement('button');
  open.className = 'mdv-btn mdv-btn-primary mdv-nav-new-btn';
  open.setAttribute('type', 'button');
  open.textContent = view.strings.newPage;
  block.appendChild(open);

  const form: HTMLElement = doc.createElement('div');
  form.className = 'mdv-nav-form mdv-hidden';
  block.appendChild(form);

  const name: HTMLInputElement = doc.createElement('input');
  name.className = 'mdv-nav-input';
  name.setAttribute('type', 'text');
  name.setAttribute('placeholder', view.strings.newPageName);
  form.appendChild(name);

  const formats: HTMLInputElement[] = buildRadioRow(doc, form, view.strings.format, 'mdv-new-format', [
    { value: 'markdown', text: '.md' },
    { value: 'html', text: '.html' }
  ]);

  const editors: HTMLInputElement[] = buildRadioRow(doc, form, view.strings.openWith, 'mdv-new-editor', [
    { value: 'sharepoint', text: view.strings.openInSharePoint },
    { value: 'app', text: view.strings.openInApp }
  ]);

  const actions: HTMLElement = doc.createElement('div');
  actions.className = 'mdv-nav-form-actions';

  const create: HTMLElement = doc.createElement('button');
  create.className = 'mdv-btn mdv-btn-primary';
  create.setAttribute('type', 'button');
  create.textContent = view.strings.create;
  create.onclick = (): void => {
    view.onCreate({
      name: name.value,
      format: formats[1].checked ? 'html' : 'markdown',
      editor: editors[1].checked ? 'app' : 'sharepoint'
    });
  };

  const cancel: HTMLElement = doc.createElement('button');
  cancel.className = 'mdv-btn';
  cancel.setAttribute('type', 'button');
  cancel.textContent = view.strings.cancel;
  cancel.onclick = (): void => {
    form.className = 'mdv-nav-form mdv-hidden';
    open.className = 'mdv-btn mdv-btn-primary mdv-nav-new-btn';
  };

  actions.appendChild(create);
  actions.appendChild(cancel);
  form.appendChild(actions);

  open.onclick = (): void => {
    form.className = 'mdv-nav-form';
    // The button stays as the heading of the open form, but stops competing
    // for attention with the create button inside it.
    open.className = 'mdv-btn mdv-nav-new-btn';
    name.focus();
  };

  return block;
}

interface IRadioChoice {
  value: string;
  text: string;
}

/** A labelled row of radio buttons; returns them in the order given. */
function buildRadioRow(
  doc: Document,
  form: HTMLElement,
  label: string,
  group: string,
  choices: IRadioChoice[]
): HTMLInputElement[] {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-nav-new-row';

  const caption: HTMLElement = doc.createElement('span');
  caption.className = 'mdv-nav-row-label';
  caption.textContent = label;
  row.appendChild(caption);

  const inputs: HTMLInputElement[] = [];

  for (let i: number = 0; i < choices.length; i++) {
    const option: HTMLElement = doc.createElement('label');
    option.className = 'mdv-nav-format';

    const input: HTMLInputElement = doc.createElement('input');
    input.setAttribute('type', 'radio');
    input.setAttribute('name', group);
    input.value = choices[i].value;
    input.checked = i === 0;
    inputs.push(input);

    const text: HTMLElement = doc.createElement('span');
    text.textContent = choices[i].text;

    option.appendChild(input);
    option.appendChild(text);
    row.appendChild(option);
  }

  form.appendChild(row);
  return inputs;
}

function buildTree(doc: Document, view: INavPanelView): HTMLElement {
  const block: HTMLElement = doc.createElement('div');
  block.className = 'mdv-nav-tree';

  const title: HTMLElement = doc.createElement('div');
  title.className = 'mdv-nav-title';
  title.textContent = view.strings.location;
  block.appendChild(title);

  // The path down to the current folder, one step per level, so the reader can
  // climb back out; then what sits in the folder they are in.
  const trail: INavEntry[] = [{ path: '', name: view.strings.root }].concat(view.ancestors);

  for (let i: number = 0; i < trail.length; i++) {
    block.appendChild(buildFolderRow(doc, view, trail[i], i, i === trail.length - 1));
  }

  const depth: number = trail.length;

  for (let i: number = 0; i < view.folders.length; i++) {
    block.appendChild(buildFolderRow(doc, view, view.folders[i], depth, false));
  }

  if (view.files.length === 0) {
    const empty: HTMLElement = doc.createElement('div');
    empty.className = 'mdv-nav-empty';
    empty.textContent = view.strings.empty;
    block.appendChild(empty);
    return block;
  }

  for (let i: number = 0; i < view.files.length; i++) {
    block.appendChild(buildFileRow(doc, view, view.files[i], depth));
  }

  return block;
}

function buildFolderRow(
  doc: Document,
  view: INavPanelView,
  entry: INavEntry,
  depth: number,
  current: boolean
): HTMLElement {
  const row: HTMLElement = doc.createElement('button');
  row.className = current ? 'mdv-nav-row mdv-nav-folder mdv-nav-current' : 'mdv-nav-row mdv-nav-folder';
  row.setAttribute('type', 'button');
  row.style.paddingLeft = 8 + depth * 14 + 'px';
  row.textContent = entry.name;
  row.onclick = (): void => view.onFolder(entry.path);
  return row;
}

function buildFileRow(doc: Document, view: INavPanelView, entry: INavEntry, depth: number): HTMLElement {
  const current: boolean = entry.path.toLowerCase() === view.currentPath.toLowerCase();
  const row: HTMLElement = doc.createElement('a');

  row.className = current ? 'mdv-nav-row mdv-nav-file mdv-nav-active' : 'mdv-nav-row mdv-nav-file';
  row.setAttribute('href', view.buildDocUrl(entry.path));
  row.style.paddingLeft = 8 + depth * 14 + 'px';
  row.textContent = entry.name;
  row.onclick = (event: MouseEvent): void => {
    if (event.defaultPrevented || event.button !== 0) {
      return;
    }
    if (event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) {
      return;
    }
    event.preventDefault();
    view.onOpen(entry.path);
  };

  return row;
}

export interface IRenamePanelStrings {
  title: string;
  /** Label of the folder box. */
  folder: string;
  /** Label of the file name box. */
  name: string;
  /** Hint under the boxes. */
  hint: string;
  save: string;
  cancel: string;
}

export interface IRenamePanelView {
  container: HTMLElement;
  /** Library relative path of the document being renamed. */
  path: string;
  /**
   * Folders that already exist in the library, offered as suggestions. Moving
   * to a folder that does not exist fails, so the ones that do are worth
   * showing rather than making people remember them.
   */
  folders?: string[];
  strings: IRenamePanelStrings;
  /** Receives the new library relative path; the host performs the move. */
  onSave(path: string): void;
  onCancel(): void;
}

/**
 * Rename and move in one panel: the folder and the file name are edited
 * separately, and joined back into a library relative path on save.
 */
export function renderRenamePanel(view: IRenamePanelView): void {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  const separator: number = view.path.lastIndexOf('/');
  const folderValue: string = separator < 0 ? '' : view.path.slice(0, separator);
  const nameValue: string = separator < 0 ? view.path : view.path.slice(separator + 1);

  const panel: HTMLElement = doc.createElement('div');
  panel.className = 'mdv-rename-panel';

  const title: HTMLElement = doc.createElement('div');
  title.className = 'mdv-tag-panel-title';
  title.textContent = view.strings.title;
  panel.appendChild(title);

  const body: HTMLElement = doc.createElement('div');
  body.className = 'mdv-rename-body';

  const folder: HTMLInputElement = buildLabelledInput(doc, body, view.strings.folder, folderValue);
  const name: HTMLInputElement = buildLabelledInput(doc, body, view.strings.name, nameValue);

  if (view.folders && view.folders.length > 0) {
    // A datalist suggests while still allowing a folder to be typed, which a
    // plain dropdown would not.
    const listId: string = 'mdv-folder-list';
    const list: HTMLElement = doc.createElement('datalist');
    list.id = listId;

    // The library root first: it is where a file goes when the box is emptied.
    const all: string[] = [''].concat(view.folders);
    for (let i: number = 0; i < all.length; i++) {
      const option: HTMLElement = doc.createElement('option');
      option.setAttribute('value', all[i]);
      list.appendChild(option);
    }

    body.appendChild(list);
    folder.setAttribute('list', listId);
  }

  const hint: HTMLElement = doc.createElement('div');
  hint.className = 'mdv-tag-hint';
  hint.textContent = view.strings.hint;
  body.appendChild(hint);

  panel.appendChild(body);

  const footer: HTMLElement = doc.createElement('div');
  footer.className = 'mdv-tag-panel-footer';

  const save: HTMLElement = doc.createElement('button');
  save.className = 'mdv-btn mdv-btn-primary';
  save.setAttribute('type', 'button');
  save.textContent = view.strings.save;
  save.onclick = (): void => view.onSave(joinFolderAndName(folder.value, name.value));

  const cancel: HTMLElement = doc.createElement('button');
  cancel.className = 'mdv-btn';
  cancel.setAttribute('type', 'button');
  cancel.textContent = view.strings.cancel;
  cancel.onclick = (): void => view.onCancel();

  footer.appendChild(save);
  footer.appendChild(cancel);
  panel.appendChild(footer);

  view.container.appendChild(panel);
}

/**
 * The folders between the library root and `folder`, outermost first, each with
 * the path that leads to it. `guide/api` becomes `guide` then `guide/api`.
 */
export function buildAncestors(folder: string): INavEntry[] {
  const normalized: string = normalizePath(folder);
  if (normalized === '') {
    return [];
  }

  const segments: string[] = normalized.split('/');
  const entries: INavEntry[] = [];
  let walked: string = '';

  for (let i: number = 0; i < segments.length; i++) {
    walked = walked === '' ? segments[i] : walked + '/' + segments[i];
    entries.push({ name: segments[i], path: walked });
  }

  return entries;
}

/**
 * The file name to create: the typed name when it already carries an extension
 * the viewer renders, otherwise the one the chosen format implies. Returns `''`
 * when the name is unusable, so the host can say so instead of creating
 * something surprising.
 */
export function withExtension(name: string, format: DocumentFormat): string {
  const trimmed: string = name.replace(/^\s+|\s+$/g, '').replace(/^[/\\]+|[/\\]+$/g, '');

  if (trimmed === '' || trimmed.indexOf('/') >= 0 || trimmed.indexOf('\\') >= 0) {
    return '';
  }

  return isDocumentPath(trimmed) ? trimmed : trimmed + (format === 'html' ? '.html' : '.md');
}

/** Joins the two boxes back into a library relative path. */
export function joinFolderAndName(folder: string, name: string): string {
  const left: string = trimPath(folder);
  const right: string = trimPath(name);

  if (right === '') {
    return '';
  }

  return left === '' ? right : left + '/' + right;
}

/** Back slashes to slashes, then space and stray separators off both ends. */
function trimPath(value: string): string {
  return value
    .replace(/\\/g, '/')
    .replace(/^\s+|\s+$/g, '')
    .replace(/^\/+|\/+$/g, '');
}

function buildLabelledInput(
  doc: Document,
  body: HTMLElement,
  label: string,
  value: string
): HTMLInputElement {
  const row: HTMLElement = doc.createElement('label');
  row.className = 'mdv-rename-row';

  const text: HTMLElement = doc.createElement('span');
  text.className = 'mdv-rename-label';
  text.textContent = label;
  row.appendChild(text);

  const input: HTMLInputElement = doc.createElement('input');
  input.className = 'mdv-tag-input';
  input.setAttribute('type', 'text');
  input.value = value;
  row.appendChild(input);

  body.appendChild(row);
  return input;
}
