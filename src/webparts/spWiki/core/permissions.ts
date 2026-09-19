/**
 * Limiting who can see a folder or a file.
 *
 * The whole feature turns on one SharePoint concept: a folder or file inherits
 * its permissions from its parent until somebody *breaks* that inheritance, at
 * which point it keeps a copy of its own and the parent stops reaching it.
 *
 * Two consequences shape everything here.
 *
 * 1. Restricting a folder restricts everything under it — except any descendant
 *    that already broke its own inheritance. Those are invisible holes in a
 *    restriction, so {@link IPermissionState.brokenDescendants} exists and the
 *    panel always shows it.
 * 2. Breaking inheritance is not a small step. It is confirmed explicitly, the
 *    way back is offered next to it, and the break copies the current
 *    permissions so that no one is locked out mid-edit.
 *
 * Everything in this file is pure DOM over data handed in, so it can be tested
 * without a tenant. The REST calls live in the service.
 */

/** Settle time before a keystroke turns into a directory lookup. */
const SUGGEST_DELAY_MS: number = 300;

/** Shortest query worth sending; below this everything matches. */
const SUGGEST_MIN_LENGTH: number = 2;

/** What kind of thing a permission was granted to. */
export type PrincipalKind = 'sharePointGroup' | 'user' | 'securityGroup';

export interface IPrincipal {
  /**
   * SharePoint's own id for the principal. 0 for a directory result that has
   * never been used on this site — the host resolves it before granting.
   */
  id: number;
  /** Claims login name, which is what resolves a directory principal. */
  loginName: string;
  title: string;
  kind: PrincipalKind;
  email: string;
}

export interface IRoleDefinition {
  id: number;
  name: string;
}

/** One principal and what it is allowed to do here. */
export interface IAssignment {
  principal: IPrincipal;
  roles: IRoleDefinition[];
}

export interface IPermissionState {
  /** Library relative path this is about. */
  path: string;
  isFolder: boolean;
  /** False while the item still inherits from its parent. */
  hasUnique: boolean;
  assignments: IAssignment[];
  /**
   * Descendants holding permissions of their own. A change made here does not
   * reach them, which is the one way a restriction silently leaks.
   */
  brokenDescendants: string[];
}

/**
 * Assignments worth showing.
 *
 * SharePoint keeps assignments with no role bindings left, and the site
 * collection administrators entry, neither of which anyone can act on here.
 */
export function visibleAssignments(assignments: IAssignment[]): IAssignment[] {
  const visible: IAssignment[] = [];

  for (let i: number = 0; i < assignments.length; i++) {
    if (assignments[i].roles.length === 0) {
      continue;
    }

    // "Limited Access" is SharePoint's bookkeeping for a parent folder someone
    // can traverse; it grants nothing and only confuses the list.
    if (assignments[i].roles.length === 1 && isLimitedAccess(assignments[i].roles[0])) {
      continue;
    }

    visible.push(assignments[i]);
  }

  return visible;
}

/**
 * Where the read-only level sits in the list, or 0 when there is none.
 *
 * Matched by name because the ids are per-site. The English and Japanese names
 * are both checked, since the site language decides which one comes back.
 */
export function indexOfReadRole(roles: IRoleDefinition[]): number {
  for (let i: number = 0; i < roles.length; i++) {
    const name: string = roles[i].name.toLowerCase();
    if (name === 'read' || name === '閲覧') {
      return i;
    }
  }
  return 0;
}

function isLimitedAccess(role: IRoleDefinition): boolean {
  const name: string = role.name.toLowerCase();
  return name === 'limited access' || name === '制限付きアクセス';
}

/** Roles a person can be granted here, in the order they should be offered. */
export function grantableRoles(roles: IRoleDefinition[]): IRoleDefinition[] {
  const grantable: IRoleDefinition[] = [];

  for (let i: number = 0; i < roles.length; i++) {
    // Limited Access is assigned by SharePoint, never chosen; Full Control is
    // not something a wiki page should be handing out.
    const name: string = roles[i].name.toLowerCase();
    if (isLimitedAccess(roles[i]) || name === 'full control' || name === 'フル コントロール') {
      continue;
    }
    grantable.push(roles[i]);
  }

  return grantable;
}

/**
 * What an inherited item is inheriting *from*, in words.
 *
 * "Inherits its permissions" leaves the reader asking whose. Naming the parent
 * answers it: the folder above, or the team itself once there is no folder
 * left. `inheritedFrom` carries a `{0}` for the name.
 */
export function parentLabel(path: string, teamLabel: string, folderSuffix: string): string {
  const separator: number = path.replace(/\/+$/, '').lastIndexOf('/');
  if (separator < 0) {
    return teamLabel;
  }

  const parent: string = path.slice(0, separator);
  const name: string = parent.slice(parent.lastIndexOf('/') + 1);

  return name === '' ? teamLabel : name + folderSuffix;
}

export interface IPermissionStrings {
  title: string;
  /** Badge for an item that follows its parent; `{0}` is the parent's name. */
  inheritedFrom: string;
  /** Badge for an item with permissions of its own. */
  unique: string;
  /** Stands in for the parent when the item sits at the top of the library. */
  team: string;
  /** Appended to a folder name in the badge, e.g. ` フォルダー`. */
  folderSuffix: string;
  /** Heading above the list of who has access. */
  whoHasAccess: string;
  /** Button that breaks inheritance. */
  restrict: string;
  /** Button that returns to inheritance. */
  restore: string;
  /** Heading of the confirmation block. */
  confirmTitle: string;
  /** Body of the confirmation for breaking inheritance. `{0}` is the path. */
  confirmRestrict: string;
  /** Body of the confirmation for restoring inheritance. `{0}` is the path. */
  confirmRestore: string;
  confirmYes: string;
  confirmNo: string;
  /** Label of the box that finds a group or a person. */
  addLabel: string;
  addPlaceholder: string;
  /** Button that runs the directory search. */
  search: string;
  /** Prefix of the message shown when the directory search failed. */
  searchFailed: string;
  /** Button that grants the selected principal the selected role. */
  add: string;
  /** Button that takes a principal's access away. */
  remove: string;
  searching: string;
  noResults: string;
  /** Heading of the list of descendants with their own permissions. */
  brokenTitle: string;
  /** Explains why that list matters. */
  brokenIntro: string;
  /** Warning shown when the target is a single file. */
  fileScopeWarning: string;
  /** Shown while the state is being read. */
  loading: string;
  close: string;
}

export interface IPermissionView {
  container: HTMLElement;
  state: IPermissionState;
  /** Roles offered when granting access. */
  roles: IRoleDefinition[];
  strings: IPermissionStrings;
  /** Finds groups and people matching what was typed. */
  searchPrincipals(query: string): Promise<IPrincipal[]>;
  onRestrict(): void;
  onRestore(): void;
  onAdd(principal: IPrincipal, roleId: number): void;
  onRemove(principal: IPrincipal): void;
  onClose(): void;
}

/** Lets the host report a failure into the panel. */
export interface IPermissionHandle {
  /** Shows a message under the buttons; empty clears it. */
  setError(message: string): void;
}

/**
 * The permissions panel.
 *
 * It always shows the state first — inherited or not, who has access, what is
 * out of reach below — and only then the actions. Someone about to restrict a
 * folder needs to know what they are starting from.
 */
export function renderPermissionPanel(view: IPermissionView): IPermissionHandle {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  const panel: HTMLElement = doc.createElement('div');
  panel.className = 'mdv-perm-panel';

  // Heading, state and the action all on one line: the state is a label, not a
  // paragraph, and the thing that changes it belongs next to it.
  const heading: HTMLElement = doc.createElement('div');
  heading.className = 'mdv-perm-heading';

  const title: HTMLElement = doc.createElement('span');
  title.className = 'mdv-tag-panel-title mdv-perm-heading-title';
  title.textContent = view.strings.title;
  heading.appendChild(title);

  const badge: HTMLElement = doc.createElement('span');
  badge.className = view.state.hasUnique
    ? 'mdv-perm-badge mdv-perm-badge-unique'
    : 'mdv-perm-badge';
  badge.textContent = view.state.hasUnique
    ? view.strings.unique
    : view.strings.inheritedFrom.replace(
        '{0}',
        parentLabel(view.state.path, view.strings.team, view.strings.folderSuffix)
      );
  heading.appendChild(badge);

  const confirmHost: HTMLElement = doc.createElement('div');
  confirmHost.className = 'mdv-perm-confirm-host';

  const action: HTMLButtonElement = doc.createElement('button');
  action.className = 'mdv-btn mdv-btn-primary mdv-perm-action';
  action.setAttribute('type', 'button');
  action.textContent = view.state.hasUnique ? view.strings.restore : view.strings.restrict;
  action.onclick = (): void => {
    renderConfirm(doc, confirmHost, view, action);
  };
  heading.appendChild(action);

  panel.appendChild(heading);
  panel.appendChild(confirmHost);

  panel.appendChild(buildAccessList(doc, view));

  if (view.state.hasUnique) {
    panel.appendChild(buildAddRow(doc, view));
  }

  if (view.state.brokenDescendants.length > 0) {
    panel.appendChild(buildBrokenList(doc, view));
  }

  const footer: HTMLElement = doc.createElement('div');
  footer.className = 'mdv-tag-panel-footer';

  const close: HTMLButtonElement = doc.createElement('button');
  close.className = 'mdv-btn';
  close.setAttribute('type', 'button');
  close.textContent = view.strings.close;
  close.onclick = (): void => view.onClose();
  footer.appendChild(close);

  panel.appendChild(footer);

  const error: HTMLElement = doc.createElement('div');
  error.className = 'mdv-tag-error';
  error.setAttribute('role', 'alert');
  panel.appendChild(error);

  view.container.appendChild(panel);

  return {
    setError: (message: string): void => {
      error.textContent = message;
    }
  };
}

/**
 * The confirmation, in the panel rather than a browser dialog.
 *
 * It names the path and says what will happen in words, because "are you sure?"
 * on its own is a question nobody can answer correctly.
 */
function renderConfirm(
  doc: Document,
  host: HTMLElement,
  view: IPermissionView,
  trigger: HTMLButtonElement
): void {
  host.innerHTML = '';
  trigger.disabled = true;

  const block: HTMLElement = doc.createElement('div');
  block.className = 'mdv-perm-confirm';
  block.setAttribute('role', 'alertdialog');

  const title: HTMLElement = doc.createElement('div');
  title.className = 'mdv-perm-confirm-title';
  title.textContent = view.strings.confirmTitle;
  block.appendChild(title);

  const body: HTMLElement = doc.createElement('p');
  body.className = 'mdv-perm-confirm-body';
  body.textContent = (view.state.hasUnique
    ? view.strings.confirmRestore
    : view.strings.confirmRestrict
  ).replace('{0}', view.state.path);
  block.appendChild(body);

  // The cost of per-file permissions belongs here, at the moment somebody is
  // about to create one — not on every visit to the panel, where it is just
  // noise about a decision nobody is making.
  if (!view.state.isFolder && !view.state.hasUnique) {
    block.appendChild(warning(doc, view.strings.fileScopeWarning));
  }

  const actions: HTMLElement = doc.createElement('div');
  actions.className = 'mdv-perm-confirm-actions';

  const yes: HTMLButtonElement = doc.createElement('button');
  yes.className = 'mdv-btn mdv-btn-primary';
  yes.setAttribute('type', 'button');
  yes.textContent = view.strings.confirmYes;
  yes.onclick = (): void => {
    host.innerHTML = '';
    trigger.disabled = false;
    if (view.state.hasUnique) {
      view.onRestore();
    } else {
      view.onRestrict();
    }
  };
  actions.appendChild(yes);

  const no: HTMLButtonElement = doc.createElement('button');
  no.className = 'mdv-btn';
  no.setAttribute('type', 'button');
  no.textContent = view.strings.confirmNo;
  no.onclick = (): void => {
    host.innerHTML = '';
    trigger.disabled = false;
  };
  actions.appendChild(no);

  block.appendChild(actions);
  host.appendChild(block);

  // The safe choice is the one the keyboard lands on.
  no.focus();
}

function buildAccessList(doc: Document, view: IPermissionView): HTMLElement {
  const block: HTMLElement = doc.createElement('div');
  block.className = 'mdv-perm-block';

  const label: HTMLElement = doc.createElement('div');
  label.className = 'mdv-help-label';
  label.textContent = view.strings.whoHasAccess;
  block.appendChild(label);

  const assignments: IAssignment[] = visibleAssignments(view.state.assignments);
  const list: HTMLElement = doc.createElement('ul');
  list.className = 'mdv-perm-list';

  for (let i: number = 0; i < assignments.length; i++) {
    list.appendChild(buildAssignmentRow(doc, view, assignments[i]));
  }

  block.appendChild(list);
  return block;
}

function buildAssignmentRow(
  doc: Document,
  view: IPermissionView,
  assignment: IAssignment
): HTMLElement {
  const item: HTMLElement = doc.createElement('li');
  item.className = 'mdv-perm-row';

  const kind: HTMLElement = doc.createElement('span');
  kind.className = 'mdv-perm-kind';
  kind.setAttribute('aria-hidden', 'true');
  kind.textContent = assignment.principal.kind === 'user' ? '👤' : '👥';
  item.appendChild(kind);

  const name: HTMLElement = doc.createElement('span');
  name.className = 'mdv-perm-name';
  name.textContent = assignment.principal.title;
  item.appendChild(name);

  const roles: string[] = [];
  for (let i: number = 0; i < assignment.roles.length; i++) {
    roles.push(assignment.roles[i].name);
  }

  const role: HTMLElement = doc.createElement('span');
  role.className = 'mdv-perm-role';
  role.textContent = roles.join(' / ');
  item.appendChild(role);

  // Only removable once the item has permissions of its own; taking someone off
  // an inherited list would mean breaking inheritance behind the reader's back.
  if (view.state.hasUnique) {
    const remove: HTMLButtonElement = doc.createElement('button');
    remove.className = 'mdv-btn mdv-perm-remove';
    remove.setAttribute('type', 'button');
    remove.textContent = view.strings.remove;
    remove.onclick = (): void => view.onRemove(assignment.principal);
    item.appendChild(remove);
  }

  return item;
}

/** The box that finds a group or a person, and grants them a role. */
function buildAddRow(doc: Document, view: IPermissionView): HTMLElement {
  const block: HTMLElement = doc.createElement('div');
  block.className = 'mdv-perm-block';

  const label: HTMLElement = doc.createElement('div');
  label.className = 'mdv-help-label';
  label.textContent = view.strings.addLabel;
  block.appendChild(label);

  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-perm-add-row';

  const query: HTMLInputElement = doc.createElement('input');
  query.className = 'mdv-tag-input';
  query.setAttribute('type', 'search');
  query.setAttribute('placeholder', view.strings.addPlaceholder);
  row.appendChild(query);

  // An explicit button, because typing an address and finding no way forward is
  // exactly what happens when the only trigger is a keystroke nobody knew about.
  const searchButton: HTMLButtonElement = doc.createElement('button');
  searchButton.className = 'mdv-btn mdv-perm-search';
  searchButton.setAttribute('type', 'button');
  searchButton.textContent = view.strings.search;
  row.appendChild(searchButton);

  const roleSelect: HTMLSelectElement = doc.createElement('select');
  roleSelect.className = 'mdv-perm-role-select';

  const roles: IRoleDefinition[] = grantableRoles(view.roles);
  const defaultIndex: number = indexOfReadRole(roles);

  for (let i: number = 0; i < roles.length; i++) {
    const option: HTMLOptionElement = doc.createElement('option');
    option.value = String(roles[i].id);
    option.textContent = roles[i].name;
    // Read by default: the least a grant can do, so a slip gives away less.
    option.selected = i === defaultIndex;
    roleSelect.appendChild(option);
  }
  row.appendChild(roleSelect);

  block.appendChild(row);

  const results: HTMLElement = doc.createElement('div');
  results.className = 'mdv-perm-results';
  block.appendChild(results);

  let token: number = 0;

  const run: () => void = (): void => {
    const text: string = query.value.replace(/^\s+|\s+$/g, '');
    results.innerHTML = '';

    if (text === '') {
      return;
    }

    token += 1;
    const mine: number = token;

    const busy: HTMLElement = doc.createElement('div');
    busy.className = 'mdv-related-empty';
    busy.textContent = view.strings.searching;
    results.appendChild(busy);

    view
      .searchPrincipals(text)
      .then((found: IPrincipal[]): void => {
        if (mine !== token) {
          return; // A newer search won.
        }

        results.innerHTML = '';

        if (found.length === 0) {
          const empty: HTMLElement = doc.createElement('div');
          empty.className = 'mdv-related-empty';
          empty.textContent = view.strings.noResults;
          results.appendChild(empty);
          return;
        }

        for (let i: number = 0; i < found.length; i++) {
          results.appendChild(buildResultRow(doc, view, found[i], roleSelect, results, query));
        }
      })
      .catch((error: Error): void => {
        if (mine !== token) {
          return;
        }

        // Saying why beats an empty list, which reads as "nobody matched" and
        // sends the reader looking for a person who was there all along.
        results.innerHTML = '';
        const failed: HTMLElement = doc.createElement('div');
        failed.className = 'mdv-perm-search-error';
        failed.setAttribute('role', 'alert');
        failed.textContent = view.strings.searchFailed + ' ' + String(error && error.message);
        results.appendChild(failed);
      });
  };

  searchButton.onclick = run;
  query.onkeydown = (event: KeyboardEvent): void => {
    if (event.key === 'Enter') {
      // Otherwise the browser treats it as a form submit and reloads the page.
      event.preventDefault();
      run();
    }
  };

  // Suggestions while typing, so the usual case needs no button at all. The
  // delay is there because every keystroke would otherwise be a directory
  // lookup; the button stays for anyone who would rather ask explicitly.
  const win: Window | null = doc.defaultView;
  if (win) {
    let timer: number = 0;

    query.oninput = (): void => {
      if (timer) {
        win.clearTimeout(timer);
      }

      // Two characters is where a directory search stops matching everybody.
      if (query.value.replace(/^\s+|\s+$/g, '').length < SUGGEST_MIN_LENGTH) {
        results.innerHTML = '';
        token += 1; // Abandon whatever is in flight.
        return;
      }

      timer = win.setTimeout(run, SUGGEST_DELAY_MS);
    };
  }

  return block;
}

function buildResultRow(
  doc: Document,
  view: IPermissionView,
  principal: IPrincipal,
  roleSelect: HTMLSelectElement,
  results: HTMLElement,
  query: HTMLInputElement
): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-perm-result';

  const kind: HTMLElement = doc.createElement('span');
  kind.className = 'mdv-perm-kind';
  kind.setAttribute('aria-hidden', 'true');
  kind.textContent = principal.kind === 'user' ? '👤' : '👥';
  row.appendChild(kind);

  const name: HTMLElement = doc.createElement('span');
  name.className = 'mdv-perm-name';
  name.textContent = principal.title;
  row.appendChild(name);

  if (principal.email !== '') {
    const email: HTMLElement = doc.createElement('span');
    email.className = 'mdv-perm-email';
    email.textContent = principal.email;
    row.appendChild(email);
  }

  const add: HTMLButtonElement = doc.createElement('button');
  add.className = 'mdv-btn mdv-btn-primary';
  add.setAttribute('type', 'button');
  add.textContent = view.strings.add;
  add.onclick = (): void => {
    results.innerHTML = '';
    query.value = '';
    view.onAdd(principal, Number(roleSelect.value));
  };
  row.appendChild(add);

  return row;
}

/** The descendants a change here will not reach. */
function buildBrokenList(doc: Document, view: IPermissionView): HTMLElement {
  const block: HTMLElement = doc.createElement('div');
  block.className = 'mdv-perm-block mdv-perm-broken';

  const label: HTMLElement = doc.createElement('div');
  label.className = 'mdv-help-label';
  label.textContent = view.strings.brokenTitle;
  block.appendChild(label);

  block.appendChild(warning(doc, view.strings.brokenIntro));

  const list: HTMLElement = doc.createElement('ul');
  list.className = 'mdv-perm-list';

  for (let i: number = 0; i < view.state.brokenDescendants.length; i++) {
    const item: HTMLElement = doc.createElement('li');
    item.className = 'mdv-perm-row';
    item.textContent = view.state.brokenDescendants[i];
    list.appendChild(item);
  }

  block.appendChild(list);
  return block;
}

function warning(doc: Document, message: string): HTMLElement {
  const line: HTMLElement = doc.createElement('p');
  line.className = 'mdv-perm-warning';
  line.textContent = message;
  return line;
}
