/**
 * The chrome around the rendered document: the toolbar and the tag editor.
 *
 * These builders are shared by the web part and by the local test harness, so
 * they use plain `mdv-*` class names (declared with `:global` in the web part
 * SCSS) instead of CSS-module class names, and take their labels from the
 * caller so localization stays with the host.
 */
import { buildPadlockIcon } from './icons';
import { ITagColor, tagColor } from './tagColors';
import { ITagGroup, acceptsNewTags, normalizeTags } from './tags';

/** Attribute carrying the toolbar action name. */
export const ACTION_ATTRIBUTE: string = 'data-mdv-action';

/**
 * `menu` / `home` / `path` / `edit` / `open-app` are used by the web part,
 * `full-page` / `close` by the preview dialog opened from a document library.
 */
export type ToolbarAction =
  | 'menu'
  | 'home'
  | 'path'
  | 'edit'
  | 'open-app'
  | 'permissions'
  | 'outline'
  | 'full-page'
  | 'close';

export interface IToolbarButton {
  action: ToolbarAction;
  label: string;
  /** Tooltip; falls back to the label. */
  title?: string;
  /** Rendered but not clickable, e.g. while the item metadata is loading. */
  disabled?: boolean;
}

export interface IToolbarView {
  container: HTMLElement;
  /** Library relative path of the current document. */
  path: string;
  buttons: IToolbarButton[];
  /** Tooltip of the leading hamburger button; omitted leaves it out. */
  menuLabel?: string;
  /** True while the file tree is pinned beside the document. */
  menuPinned?: boolean;
  /**
   * Name of this viewer, from the web part settings. Shown on the top row and
   * doubling as the way home; omitted leaves the row out.
   */
  title?: string;
  /** Tooltip of the title, e.g. "back to the start page". */
  titleLabel?: string;
  /** Tooltip of the file name and folder; omitted renders them as plain text. */
  pathLabel?: string;
  /** Shown in place of the folder when the document sits in the library root. */
  rootLabel?: string;
  /**
   * The padlock in front of the file name. Omitted leaves it out.
   *
   * It sits with the file name rather than among the actions on the right
   * because it says something *about* this page, and because its shape is the
   * fastest way to see whether a page is restricted — which is worth noticing
   * without having gone looking for it.
   */
  permission?: IPermissionIndicator;
  /**
   * The outline button, in front of the padlock. Omitted leaves it out.
   *
   * Left of everything else because it is about *this document* rather than
   * about the library, and because it is the control a reader reaches for most
   * often on a long page.
   */
  outline?: IOutlineIndicator;
  /**
   * Makes each folder segment a link that opens that folder as a page.
   * Both have to be given; either alone leaves the folder as plain text.
   */
  folderLinkAttribute?: string;
  buildFolderUrl?(path: string): string;
}

export interface IOutlineIndicator {
  /** Tooltip and accessible name. */
  label: string;
  /** True while the outline is pinned open beside the document. */
  pinned: boolean;
}

export interface IPermissionIndicator {
  /**
   * True when the item has permissions of its own, `undefined` while that is
   * still being read.
   *
   * It stays clickable throughout: the panel reads its own state from the path,
   * so there is nothing for it to wait for. Only the *shape* of the icon is
   * unknown for the moment the item metadata is in flight.
   */
  locked?: boolean;
  /** Tooltip and accessible name. */
  label: string;
}

/**
 * Renders the header: the menu button and the viewer's name on the top row with
 * the actions on the right, and the file name with the folder it lives in
 * underneath.
 *
 * The name is on top because it never changes — it is the thing that tells a
 * reader where they are, and the way back to the start page. The file name is
 * one row down because it is what changes as they read.
 */
export function renderToolbar(view: IToolbarView): void {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  // Own wrapper rather than styling the host element, so the host keeps its own
  // class (a CSS-module one in the web part) for the surrounding chrome.
  const bar: HTMLElement = doc.createElement('div');
  bar.className = 'mdv-toolbar';

  if (view.menuLabel) {
    const menu: HTMLElement = doc.createElement('button');
    menu.className = view.menuPinned ? 'mdv-btn mdv-menu-btn mdv-menu-btn-pinned' : 'mdv-btn mdv-menu-btn';
    menu.setAttribute('type', 'button');
    menu.setAttribute(ACTION_ATTRIBUTE, 'menu');
    menu.setAttribute('title', view.menuLabel);
    menu.setAttribute('aria-label', view.menuLabel);
    menu.setAttribute('aria-pressed', view.menuPinned ? 'true' : 'false');
    menu.appendChild(buildHamburgerIcon(doc));
    bar.appendChild(menu);
  }

  if (view.title) {
    const title: HTMLElement = doc.createElement('button');
    title.className = 'mdv-toolbar-title';
    title.setAttribute('type', 'button');
    title.setAttribute(ACTION_ATTRIBUTE, 'home');
    title.setAttribute('title', view.titleLabel || view.title);
    title.textContent = view.title;
    bar.appendChild(title);
  }

  const actions: HTMLElement = doc.createElement('span');
  actions.className = 'mdv-toolbar-actions';

  for (let i: number = 0; i < view.buttons.length; i++) {
    const button: IToolbarButton = view.buttons[i];
    const element: HTMLButtonElement = doc.createElement('button');
    element.className = 'mdv-btn';
    element.setAttribute('type', 'button');
    element.setAttribute(ACTION_ATTRIBUTE, button.action);
    element.setAttribute('title', button.title || button.label);
    element.textContent = button.label;
    if (button.disabled) {
      element.setAttribute('disabled', 'disabled');
    }
    actions.appendChild(element);
  }

  bar.appendChild(actions);
  view.container.appendChild(bar);
  view.container.appendChild(buildFileRow(doc, view));
}

/** The second row: file name, then the folder holding it. */
function buildFileRow(doc: Document, view: IToolbarView): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-toolbar-file';

  const separator: number = view.path.lastIndexOf('/');
  const folder: string = separator < 0 ? '' : view.path.slice(0, separator);
  const name: string = separator < 0 ? view.path : view.path.slice(separator + 1);

  if (view.outline) {
    row.appendChild(buildOutlineButton(doc, view.outline));
  }

  if (view.permission) {
    row.appendChild(buildPermissionButton(doc, view.permission));
  }

  // A button when it can be clicked, so keyboard users reach the rename panel
  // the same way; plain text when the host offers no rename.
  row.appendChild(buildPathPart(doc, view, 'mdv-toolbar-path', name));
  row.appendChild(buildFolderPart(doc, view, folder));

  return row;
}

/**
 * The folder, one link per segment.
 *
 * Each segment opens *that* folder rather than the whole path, so the row
 * doubles as a breadcrumb: `guide/api` gives a way into `guide` and into
 * `guide/api` separately, which is how people walk back up.
 */
function buildFolderPart(doc: Document, view: IToolbarView, folder: string): HTMLElement {
  const wrapper: HTMLElement = doc.createElement('span');
  wrapper.className = 'mdv-toolbar-folder';

  if (folder === '') {
    wrapper.textContent = view.rootLabel || '';
    return wrapper;
  }

  const linkAttribute: string | undefined = view.folderLinkAttribute;
  const buildFolderUrl: ((path: string) => string) | undefined = view.buildFolderUrl;

  if (!linkAttribute || !buildFolderUrl) {
    wrapper.textContent = folder;
    return wrapper;
  }

  const segments: string[] = folder.split('/');
  let walked: string = '';

  for (let i: number = 0; i < segments.length; i++) {
    walked = walked === '' ? segments[i] : walked + '/' + segments[i];

    if (i > 0) {
      const separator: HTMLElement = doc.createElement('span');
      separator.className = 'mdv-toolbar-folder-sep';
      separator.setAttribute('aria-hidden', 'true');
      separator.textContent = '/';
      wrapper.appendChild(separator);
    }

    const link: HTMLElement = doc.createElement('a');
    link.className = 'mdv-toolbar-folder-link';
    link.setAttribute('href', buildFolderUrl(walked));
    link.setAttribute(linkAttribute, walked);
    link.textContent = segments[i];
    wrapper.appendChild(link);
  }

  return wrapper;
}

/**
 * The outline button.
 *
 * One click shows the outline over the document, a double click pins it beside
 * the document instead. The shape stays the same either way; only the pressed
 * state says which mode it is in.
 */
function buildOutlineButton(doc: Document, outline: IOutlineIndicator): HTMLElement {
  const button: HTMLButtonElement = doc.createElement('button');

  button.className = outline.pinned
    ? 'mdv-btn mdv-outline-btn mdv-outline-btn-pinned'
    : 'mdv-btn mdv-outline-btn';
  button.setAttribute('type', 'button');
  button.setAttribute(ACTION_ATTRIBUTE, 'outline');
  button.setAttribute('title', outline.label);
  button.setAttribute('aria-label', outline.label);
  button.setAttribute('aria-pressed', outline.pinned ? 'true' : 'false');
  button.appendChild(buildOutlineIcon(doc));

  return button;
}

/** Three lines of decreasing length — a list, not a menu. */
function buildOutlineIcon(doc: Document): Element {
  const NS: string = 'http://www.w3.org/2000/svg';
  const svg: Element = doc.createElementNS(NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '14');
  svg.setAttribute('height', '14');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('focusable', 'false');

  const widths: number[] = [11, 8, 5];
  for (let i: number = 0; i < widths.length; i++) {
    const line: Element = doc.createElementNS(NS, 'rect');
    line.setAttribute('x', String(2 + i * 2));
    line.setAttribute('y', String(3 + i * 4));
    line.setAttribute('width', String(widths[i]));
    line.setAttribute('height', '1.6');
    line.setAttribute('rx', '0.8');
    line.setAttribute('fill', 'currentColor');
    svg.appendChild(line);
  }

  return svg;
}

/** The file name: a button that opens the rename panel, or plain text. */
function buildPathPart(
  doc: Document,
  view: IToolbarView,
  className: string,
  text: string
): HTMLElement {
  const part: HTMLElement = doc.createElement(view.pathLabel ? 'button' : 'span');

  part.className = view.pathLabel ? className + ' mdv-toolbar-path-action' : className;
  part.textContent = text;
  part.setAttribute('title', view.pathLabel || view.path);

  if (view.pathLabel) {
    part.setAttribute('type', 'button');
    part.setAttribute(ACTION_ATTRIBUTE, 'path');
  }

  return part;
}

/**
 * The padlock in front of the file name.
 *
 * Open or closed rather than a word: the state is worth seeing at a glance on
 * every page, and a label saying "access" says nothing about which state it is
 * in. The shape carries the meaning; the tooltip carries the detail.
 */
function buildPermissionButton(doc: Document, permission: IPermissionIndicator): HTMLElement {
  const button: HTMLButtonElement = doc.createElement('button');

  // Never disabled: greying it out until the metadata arrived made it look
  // like the button was missing, and it has nothing to wait for anyway.
  button.className =
    permission.locked === undefined
      ? 'mdv-btn mdv-perm-btn mdv-perm-btn-unknown'
      : permission.locked
        ? 'mdv-btn mdv-perm-btn mdv-perm-btn-locked'
        : 'mdv-btn mdv-perm-btn';
  button.setAttribute('type', 'button');
  button.setAttribute(ACTION_ATTRIBUTE, 'permissions');
  button.setAttribute('title', permission.label);
  button.setAttribute('aria-label', permission.label);

  button.appendChild(buildPadlockIcon(doc, permission.locked === true, 14));
  return button;
}

/** Three stacked lines, drawn inline so no icon font is needed. */
function buildHamburgerIcon(doc: Document): Element {
  const NS: string = 'http://www.w3.org/2000/svg';
  const svg: Element = doc.createElementNS(NS, 'svg');
  svg.setAttribute('viewBox', '0 0 16 16');
  svg.setAttribute('width', '15');
  svg.setAttribute('height', '15');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('focusable', 'false');

  for (let i: number = 0; i < 3; i++) {
    const line: Element = doc.createElementNS(NS, 'rect');
    line.setAttribute('x', '2');
    line.setAttribute('y', String(3 + i * 4));
    line.setAttribute('width', '12');
    line.setAttribute('height', '1.6');
    line.setAttribute('rx', '0.8');
    line.setAttribute('fill', 'currentColor');
    svg.appendChild(line);
  }

  return svg;
}

export interface ITagBarView {
  container: HTMLElement;
  /** One row per configured column; rows without tags are left out. */
  groups: ITagGroup[];
  /** Real href, so a tag can be opened in a new tab or copied. */
  buildTagUrl(tag: string): string;
  /** Plain clicks are handled in place instead of reloading the page. */
  onTag(tag: string): void;
  /**
   * Opens the tag editor for the column whose name was clicked — its internal
   * name is passed, because that is what identifies a column everywhere else.
   * Given, the column name at the start of each row becomes the way in; left
   * out, the names are plain text.
   */
  onEdit?(internalName: string): void;
  /** Tooltip of the column name when it opens the editor. */
  editLabel?: string;
}

/**
 * The tags of the current document, one row per configured column, under the
 * file name. Each tag links to the search that lists every page carrying it.
 *
 * Every configured column gets a row even when it holds no tags: the column
 * name is what opens the editor, so a document with nothing on it yet still has
 * to offer a way to put something there.
 */
export function renderTagBar(view: ITagBarView): void {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  for (let g: number = 0; g < view.groups.length; g++) {
    const group: ITagGroup = view.groups[g];

    const row: HTMLElement = doc.createElement('div');
    row.className = 'mdv-tag-row';

    // The column name is the way into the editor, so the toolbar needs no
    // button of its own: the tags are already what the reader is looking at.
    const label: HTMLElement = doc.createElement(view.onEdit ? 'button' : 'span');
    label.className = view.onEdit ? 'mdv-tag-row-label mdv-tag-row-edit' : 'mdv-tag-row-label';
    label.textContent = group.field.title;
    if (view.onEdit) {
      const onEdit: (internalName: string) => void = view.onEdit;
      const internalName: string = group.field.internalName;
      label.setAttribute('type', 'button');
      label.setAttribute('title', view.editLabel || group.field.title);
      label.onclick = (): void => onEdit(internalName);
    }
    row.appendChild(label);

    for (let i: number = 0; i < group.tags.length; i++) {
      row.appendChild(buildTagChip(doc, group.tags[i], view));
    }

    view.container.appendChild(row);
  }
}

function buildTagChip(doc: Document, tag: string, view: ITagBarView): HTMLElement {
  const colour: ITagColor = tagColor(tag);
  const chip: HTMLElement = doc.createElement('a');

  chip.className = 'mdv-tag-link';
  chip.textContent = tag;
  chip.setAttribute('href', view.buildTagUrl(tag));
  chip.style.backgroundColor = colour.background;
  chip.style.color = colour.foreground;
  chip.onclick = (event: MouseEvent): void => {
    if (event.defaultPrevented || event.button !== 0) {
      return;
    }
    // Modified clicks keep opening a new tab, like any other link.
    if (event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) {
      return;
    }
    event.preventDefault();
    view.onTag(tag);
  };

  return chip;
}

export interface ITagEditorStrings {
  title: string;
  save: string;
  cancel: string;
  /** Hint shown under a free-text column. */
  textHint: string;
  /** Shown when a choice column has no choices configured. */
  noChoices: string;
  /** Label above the box that adds a tag the column does not offer yet. */
  addLabel: string;
  /** Placeholder of that box. */
  addPlaceholder: string;
  /** Placeholder used instead while the box is disabled. */
  addDisabledPlaceholder: string;
  /** Hint under that box. */
  addHint: string;
  /**
   * Why the box is disabled. Carried by the warning icon's tooltip and by the
   * box itself rather than printed under every such column.
   */
  addDisabledHint: string;
}

export interface ITagEditorView {
  container: HTMLElement;
  /** The columns to edit — one, since the editor opens per column. */
  groups: ITagGroup[];
  strings: ITagEditorStrings;
  /** Receives the edited tags per column; the host performs the save. */
  onSave(groups: ITagGroup[]): void;
  onCancel(): void;
}

/** Lets the host talk back to an editor that is already on screen. */
export interface ITagEditorHandle {
  /**
   * Shows why the save failed, under the buttons that triggered it. An empty
   * message clears the line.
   */
  setError(message: string): void;
}

/**
 * Renders the tag editor: one section per configured column, each adapting to
 * the column type — a text box for a text column, radio buttons for a single
 * choice column, check boxes for a multi-choice one — plus a box for a tag the
 * column does not offer yet.
 */
export function renderTagEditor(view: ITagEditorView): ITagEditorHandle {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  const panel: HTMLElement = doc.createElement('div');
  panel.className = 'mdv-tag-panel';

  const heading: HTMLElement = doc.createElement('div');
  heading.className = 'mdv-tag-panel-title';
  heading.textContent = view.strings.title;
  panel.appendChild(heading);

  const readers: (() => ITagGroup)[] = [];

  for (let g: number = 0; g < view.groups.length; g++) {
    readers.push(buildTagSection(doc, panel, view.groups[g], view.strings));
  }

  const footer: HTMLElement = doc.createElement('div');
  footer.className = 'mdv-tag-panel-footer';

  // Under the buttons, because that is where the eye already is after pressing
  // one. `role="alert"` so a screen reader hears it without having to go
  // looking for what changed.
  const error: HTMLElement = doc.createElement('div');
  error.className = 'mdv-tag-error';
  error.setAttribute('role', 'alert');

  const clearError: () => void = (): void => {
    error.textContent = '';
  };

  const save: HTMLButtonElement = doc.createElement('button');
  save.className = 'mdv-btn mdv-btn-primary';
  save.setAttribute('type', 'button');
  save.textContent = view.strings.save;
  save.onclick = (): void => {
    // Cleared on the way in: a message left over from the previous attempt
    // would otherwise look like the verdict on this one.
    clearError();

    const groups: ITagGroup[] = [];
    for (let i: number = 0; i < readers.length; i++) {
      groups.push(readers[i]());
    }
    view.onSave(groups);
  };

  const cancel: HTMLButtonElement = doc.createElement('button');
  cancel.className = 'mdv-btn';
  cancel.setAttribute('type', 'button');
  cancel.textContent = view.strings.cancel;
  cancel.onclick = (): void => {
    clearError();
    view.onCancel();
  };

  footer.appendChild(save);
  footer.appendChild(cancel);
  panel.appendChild(footer);
  panel.appendChild(error);

  view.container.appendChild(panel);

  return {
    setError: (message: string): void => {
      error.textContent = message;
    }
  };
}

/**
 * One column's section of the editor. Returns a reader for what the user ended
 * up selecting, so the panel can collect every column on save.
 */
function buildTagSection(
  doc: Document,
  panel: HTMLElement,
  group: ITagGroup,
  strings: ITagEditorStrings
): () => ITagGroup {
  const section: HTMLElement = doc.createElement('div');
  section.className = 'mdv-tag-section';

  const title: HTMLElement = doc.createElement('div');
  title.className = 'mdv-tag-section-title';
  title.textContent = group.field.title;
  section.appendChild(title);

  const body: HTMLElement = doc.createElement('div');
  body.className = 'mdv-tag-panel-body';
  section.appendChild(body);

  panel.appendChild(section);

  if (group.field.kind === 'text') {
    const input: HTMLInputElement = doc.createElement('input');
    input.className = 'mdv-tag-input';
    input.setAttribute('type', 'text');
    input.value = group.tags.join('; ');
    body.appendChild(input);

    const hint: HTMLElement = doc.createElement('div');
    hint.className = 'mdv-tag-hint';
    hint.textContent = strings.textHint;
    body.appendChild(hint);

    // Normalized here so the caller always gets trimmed, de-duplicated tags,
    // whatever the user typed.
    return (): ITagGroup => {
      return { field: group.field, tags: normalizeTags(input.value.split(/[;,、，；]/)) };
    };
  }

  const multiple: boolean = group.field.kind === 'multiChoice';
  const inputs: HTMLInputElement[] = [];

  // Tags already on the document that the column does not offer as a choice —
  // someone added them by hand, or through the box below — are shown as
  // choices too, otherwise saving would silently drop them.
  const choices: string[] = group.field.choices.slice();
  for (let i: number = 0; i < group.tags.length; i++) {
    if (!isSelected(choices, group.tags[i])) {
      choices.push(group.tags[i]);
    }
  }

  if (choices.length === 0) {
    const empty: HTMLElement = doc.createElement('div');
    empty.className = 'mdv-tag-hint';
    empty.textContent = strings.noChoices;
    body.appendChild(empty);
  }

  for (let i: number = 0; i < choices.length; i++) {
    const choice: string = choices[i];
    const label: HTMLElement = doc.createElement('label');
    label.className = 'mdv-tag-option';

    const input: HTMLInputElement = doc.createElement('input');
    input.setAttribute('type', multiple ? 'checkbox' : 'radio');
    // Per column, so two columns do not fight over the same radio group.
    input.setAttribute('name', 'mdv-tag-' + group.field.internalName);
    input.value = choice;
    input.checked = isSelected(group.tags, choice);
    inputs.push(input);

    const text: HTMLElement = doc.createElement('span');
    text.textContent = choice;

    label.appendChild(input);
    label.appendChild(text);
    body.appendChild(label);
  }

  // SharePoint refuses a value outside the choices unless the column allows
  // custom values, so the box is disabled rather than letting the save fail.
  const canAdd: boolean = acceptsNewTags(group.field);

  const addLabel: HTMLElement = doc.createElement('div');
  addLabel.className = 'mdv-tag-add-label';

  const addLabelText: HTMLElement = doc.createElement('span');
  addLabelText.textContent = strings.addLabel;
  addLabel.appendChild(addLabelText);

  if (!canAdd) {
    // The reason sits on an icon instead of a line of prose: it is a
    // configuration detail of one column, and a paragraph of it under every
    // such column buried the tags the reader actually came for.
    //
    // Not the `title` attribute: the browser waits about a second before
    // showing one, and never shows it from the keyboard. This one is a real
    // element revealed by :hover and :focus, so it appears at once and a
    // keyboard reader can reach it.
    addLabel.appendChild(buildTooltip(doc, '!', strings.addDisabledHint, 'mdv-tag-warn'));
  }

  body.appendChild(addLabel);

  const extra: HTMLInputElement = doc.createElement('input');
  extra.className = canAdd
    ? 'mdv-tag-input mdv-tag-add'
    : 'mdv-tag-input mdv-tag-add mdv-tag-add-disabled';
  extra.setAttribute('type', 'text');

  if (canAdd) {
    extra.setAttribute('placeholder', strings.addPlaceholder);
  } else {
    // A greyed box on its own reads as "nothing to type yet". Saying so in the
    // placeholder is what makes it read as "you cannot type here", without
    // waiting for a hover over the icon.
    extra.setAttribute('disabled', 'disabled');
    extra.setAttribute('placeholder', strings.addDisabledPlaceholder);
    extra.setAttribute('title', strings.addDisabledHint);
  }

  body.appendChild(extra);

  if (canAdd) {
    const addHint: HTMLElement = doc.createElement('div');
    addHint.className = 'mdv-tag-hint';
    addHint.textContent = strings.addHint;
    body.appendChild(addHint);
  }

  return (): ITagGroup => {
    const selected: string[] = [];

    for (let i: number = 0; i < inputs.length; i++) {
      if (inputs[i].checked) {
        selected.push(inputs[i].value);
      }
    }

    if (canAdd) {
      const added: string[] = normalizeTags(extra.value.split(/[;,、，；]/));
      for (let i: number = 0; i < added.length; i++) {
        selected.push(added[i]);
      }
    }

    return { field: group.field, tags: normalizeTags(selected) };
  };
}

/**
 * A marker carrying an explanation that appears the moment it is hovered or
 * focused, in place of the `title` attribute's second-long delay.
 *
 * `tabindex="0"` is what puts it on the keyboard path; the bubble is a child so
 * plain CSS (`:hover`, `:focus-within`) can reveal it with no script at all.
 */
export function buildTooltip(
  doc: Document,
  marker: string,
  text: string,
  markerClass: string
): HTMLElement {
  const host: HTMLElement = doc.createElement('span');
  host.className = 'mdv-tip-host';
  host.setAttribute('tabindex', '0');
  host.setAttribute('role', 'note');
  host.setAttribute('aria-label', text);

  const icon: HTMLElement = doc.createElement('span');
  icon.className = markerClass;
  icon.setAttribute('aria-hidden', 'true');
  icon.textContent = marker;
  host.appendChild(icon);

  const bubble: HTMLElement = doc.createElement('span');
  bubble.className = 'mdv-tip';
  bubble.textContent = text;
  host.appendChild(bubble);

  return host;
}

function isSelected(tags: string[], choice: string): boolean {
  for (let i: number = 0; i < tags.length; i++) {
    if (tags[i].toLowerCase() === choice.toLowerCase()) {
      return true;
    }
  }
  return false;
}

export interface IToastView {
  /** Host the toast is appended to; it positions itself against the viewport. */
  container: HTMLElement;
  message: string;
  /** How long it stays fully visible, in milliseconds. */
  durationMs?: number;
}

/** Visible time before the fade starts, when the caller does not say. */
const TOAST_DURATION_MS: number = 1400;

/** Length of the fade, kept in step with the transition in the stylesheet. */
const TOAST_FADE_MS: number = 300;

/**
 * A confirmation in the bottom right corner that fades out by itself.
 *
 * Anything the reader has to read carefully or act on stays in the strip under
 * the toolbar ({@link renderNotice}). This is for the other kind — "copied",
 * "saved" — which used to push the document down and sit there until dismissed,
 * for news that stopped being news the moment it appeared.
 */
export function showToast(view: IToastView): void {
  const doc: Document = view.container.ownerDocument;

  const toast: HTMLElement = doc.createElement('div');
  toast.className = 'mdv-toast';
  // Polite: it is a confirmation, so it must not interrupt what is being read.
  toast.setAttribute('role', 'status');
  toast.textContent = view.message;
  view.container.appendChild(toast);

  const visible: number = view.durationMs === undefined ? TOAST_DURATION_MS : view.durationMs;

  const win: Window | null = doc.defaultView;
  if (!win) {
    return;
  }

  win.setTimeout((): void => {
    toast.className = 'mdv-toast mdv-toast-leaving';

    win.setTimeout((): void => {
      if (toast.parentNode) {
        toast.parentNode.removeChild(toast);
      }
    }, TOAST_FADE_MS);
  }, visible);
}

export interface INoticeView {
  container: HTMLElement;
  message: string;
  /** Optional monospace detail, such as a file path. */
  detail?: string;
  /** Label of the copy button; omitted when there is nothing to copy. */
  copyLabel?: string;
  onCopy?(): void;
  onClose(): void;
  closeLabel: string;
}

/**
 * Small dismissible strip under the toolbar. Used for the "open in app"
 * fallback and for save confirmations.
 */
export function renderNotice(view: INoticeView): void {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  const notice: HTMLElement = doc.createElement('div');
  notice.className = 'mdv-notice';

  const message: HTMLElement = doc.createElement('div');
  message.className = 'mdv-notice-message';
  message.textContent = view.message;
  notice.appendChild(message);

  if (view.detail) {
    const detail: HTMLElement = doc.createElement('code');
    detail.className = 'mdv-notice-detail';
    detail.textContent = view.detail;
    notice.appendChild(detail);
  }

  const actions: HTMLElement = doc.createElement('div');
  actions.className = 'mdv-notice-actions';

  if (view.copyLabel && view.onCopy) {
    const copy: HTMLButtonElement = doc.createElement('button');
    copy.className = 'mdv-btn';
    copy.setAttribute('type', 'button');
    copy.textContent = view.copyLabel;
    copy.onclick = (): void => {
      if (view.onCopy) {
        view.onCopy();
      }
    };
    actions.appendChild(copy);
  }

  const close: HTMLButtonElement = doc.createElement('button');
  close.className = 'mdv-btn';
  close.setAttribute('type', 'button');
  close.textContent = view.closeLabel;
  close.onclick = (): void => view.onClose();
  actions.appendChild(close);

  notice.appendChild(actions);
  view.container.appendChild(notice);
}
