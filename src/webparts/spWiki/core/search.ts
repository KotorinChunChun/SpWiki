/**
 * Searching the library: by page name, and through the text of the documents.
 *
 * The two are separated because they cost very different things. A name search
 * runs over the listing the viewer already holds and is instant. A body search
 * has to read the documents, so it is one request per file — bounded, reported
 * and cancellable rather than pretended to be free.
 *
 * Everything here is pure: the matching, the snippet extraction and the DOM.
 * Fetching belongs to the host, which is the only part that needs SharePoint.
 */

import { ITreeFolder, buildTree, isEmptyTree, renderTree } from './pageTree';
import { dirname, normalizePath } from './pathUtils';
import { IRelatedPage } from './relatedPages';

/** Which of the two searches to run. */
export type SearchMode = 'name' | 'content';

/** One hit, with the line that matched when the body was searched. */
export interface ISearchHit {
  page: IRelatedPage;
  /** Text around the first match; empty for a name search. */
  snippet: string;
}

/** Trims whitespace. JavaScript's \s covers U+3000 IDEOGRAPHIC SPACE too. */
function trim(value: string): string {
  return value.replace(/^\s+|\s+$/g, '');
}

/**
 * The documents a search covers.
 *
 * An empty folder means the whole library. The comparison is on the path
 * prefix, so `guide` takes `guide/api/rest.md` with it — a scope that stopped
 * at the first level would be surprising for a wiki, where the sub-folders are
 * the whole point.
 */
export function pagesInScope(pages: IRelatedPage[], folder: string): IRelatedPage[] {
  const scope: string = normalizePath(trim(folder).replace(/^\/+|\/+$/g, ''));
  if (scope === '') {
    return pages.slice();
  }

  const prefix: string = scope.toLowerCase() + '/';
  const matches: IRelatedPage[] = [];

  for (let i: number = 0; i < pages.length; i++) {
    if (pages[i].path.toLowerCase().slice(0, prefix.length) === prefix) {
      matches.push(pages[i]);
    }
  }

  return matches;
}

/**
 * Pages whose name or folder contains the term.
 *
 * The folder counts, so searching `guide` finds everything filed under it even
 * when no file is called that — which is how people look for a section.
 */
export function matchByName(pages: IRelatedPage[], term: string): ISearchHit[] {
  const wanted: string = trim(term).toLowerCase();
  if (wanted === '') {
    return [];
  }

  const hits: ISearchHit[] = [];

  for (let i: number = 0; i < pages.length; i++) {
    if (pages[i].path.toLowerCase().indexOf(wanted) >= 0) {
      hits.push({ page: pages[i], snippet: '' });
    }
  }

  return hits;
}

/**
 * Only the pages holding permissions of their own.
 *
 * Reviewing what has been restricted is a question on its own — "show me
 * everything that is not open to the team" — so it works with an empty search
 * term, where every other filter would return nothing.
 */
export function onlyRestricted(pages: IRelatedPage[]): IRelatedPage[] {
  const restricted: IRelatedPage[] = [];

  for (let i: number = 0; i < pages.length; i++) {
    if (pages[i].hasUniquePermissions === true) {
      restricted.push(pages[i]);
    }
  }

  return restricted;
}

/** Number of characters kept either side of a body match. */
const SNIPPET_MARGIN: number = 40;

/**
 * The text around the first occurrence of the term, on one line.
 *
 * Returns '' when the term is not there, which is what makes this double as the
 * match test for a body search.
 */
export function extractSnippet(text: string, term: string): string {
  const wanted: string = trim(term).toLowerCase();
  if (wanted === '') {
    return '';
  }

  const index: number = text.toLowerCase().indexOf(wanted);
  if (index < 0) {
    return '';
  }

  const start: number = index - SNIPPET_MARGIN < 0 ? 0 : index - SNIPPET_MARGIN;
  const end: number = index + wanted.length + SNIPPET_MARGIN;

  // Newlines and runs of spaces collapse: a snippet is one line by definition,
  // and Markdown is full of both.
  const raw: string = text.slice(start, end).replace(/\s+/g, ' ');

  return (start > 0 ? '…' : '') + trim(raw) + (end < text.length ? '…' : '');
}

export interface ISearchStrings {
  /** Heading above the results. */
  title: string;
  termPlaceholder: string;
  /** Labels of the two modes. */
  modeName: string;
  modeContent: string;
  /** Label and placeholder of the folder box. */
  scope: string;
  scopePlaceholder: string;
  /** Label of the submit button. */
  submit: string;
  /** Shown while a body search is reading the documents. */
  searching: string;
  /** `{0}` is replaced with the number of hits. */
  resultCount: string;
  empty: string;
  /** Shown when the box is empty. */
  prompt: string;
  /** Label of the restricted-only filter. */
  restrictedOnly: string;
  /** Tooltip of the padlock beside a restricted result. */
  locked: string;
}

export interface ISearchRequest {
  term: string;
  mode: SearchMode;
  /** Library relative folder; '' is the whole library. */
  scope: string;
  /** Narrows the results to pages holding permissions of their own. */
  restrictedOnly?: boolean;
}

export interface ISearchView {
  container: HTMLElement;
  request: ISearchRequest;
  strings: ISearchStrings;
  /** Address of a document, so every hit is a real link. */
  buildDocUrl(path: string): string;
  /** Attribute the host's click handling looks for on an internal link. */
  linkAttribute: string;
  /** Runs the search; the host fetches and calls back into {@link ISearchHandle}. */
  onSearch(request: ISearchRequest): void;
}

/** Lets the host fill in the results once it has them. */
export interface ISearchHandle {
  /** Replaces the result area with a progress line. */
  setBusy(message: string): void;
  /** Replaces it with the hits, grouped into the folder tree they came from. */
  setResults(hits: ISearchHit[]): void;
}

/**
 * The search page: the form, then the results.
 *
 * The form is rendered once and left alone while a search runs, so what was
 * typed stays on screen and can be adjusted — a page that blanks its own form
 * to show a spinner makes refining a query needlessly hostile.
 */
export function renderSearchPage(view: ISearchView): ISearchHandle {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  const heading: HTMLElement = doc.createElement('h1');
  heading.textContent = view.strings.title;
  view.container.appendChild(heading);

  const form: HTMLElement = doc.createElement('div');
  form.className = 'mdv-search-form';

  const term: HTMLInputElement = doc.createElement('input');
  term.className = 'mdv-tag-input mdv-search-term';
  term.setAttribute('type', 'search');
  term.setAttribute('placeholder', view.strings.termPlaceholder);
  term.value = view.request.term;
  form.appendChild(term);

  const scope: HTMLInputElement = doc.createElement('input');
  scope.className = 'mdv-tag-input mdv-search-scope';
  scope.setAttribute('type', 'text');
  scope.setAttribute('placeholder', view.strings.scopePlaceholder);
  scope.setAttribute('aria-label', view.strings.scope);
  scope.value = view.request.scope;
  form.appendChild(scope);

  const submit: HTMLButtonElement = doc.createElement('button');
  submit.className = 'mdv-btn mdv-btn-primary';
  submit.setAttribute('type', 'button');
  submit.textContent = view.strings.submit;
  form.appendChild(submit);

  const modes: HTMLElement = doc.createElement('div');
  modes.className = 'mdv-search-modes';

  const nameMode: HTMLInputElement = buildMode(
    doc,
    modes,
    view.strings.modeName,
    view.request.mode === 'name'
  );
  const contentMode: HTMLInputElement = buildMode(
    doc,
    modes,
    view.strings.modeContent,
    view.request.mode === 'content'
  );

  const restrictedLabel: HTMLElement = doc.createElement('label');
  restrictedLabel.className = 'mdv-tag-option mdv-search-restricted';

  const restricted: HTMLInputElement = doc.createElement('input');
  restricted.setAttribute('type', 'checkbox');
  restricted.checked = view.request.restrictedOnly === true;
  restrictedLabel.appendChild(restricted);

  const restrictedText: HTMLElement = doc.createElement('span');
  restrictedText.textContent = view.strings.restrictedOnly;
  restrictedLabel.appendChild(restrictedText);

  modes.appendChild(restrictedLabel);

  form.appendChild(modes);
  view.container.appendChild(form);

  const results: HTMLElement = doc.createElement('div');
  results.className = 'mdv-search-results';
  view.container.appendChild(results);

  const run: () => void = (): void => {
    view.onSearch({
      term: trim(term.value),
      mode: contentMode.checked ? 'content' : 'name',
      scope: trim(scope.value),
      restrictedOnly: restricted.checked
    });
  };

  submit.onclick = run;
  term.onkeydown = (event: KeyboardEvent): void => {
    if (event.key === 'Enter') {
      run();
    }
  };
  scope.onkeydown = term.onkeydown;
  // Switching what to search means the shown results are about the other one.
  nameMode.onchange = run;
  contentMode.onchange = run;
  restricted.onchange = run;

  const handle: ISearchHandle = {
    setBusy: (message: string): void => {
      results.innerHTML = '';
      const line: HTMLElement = doc.createElement('p');
      line.className = 'mdv-related-empty';
      line.setAttribute('role', 'status');
      line.textContent = message;
      results.appendChild(line);
    },
    setResults: (hits: ISearchHit[]): void => {
      renderResults(doc, results, hits, view);
    }
  };

  // "Show me everything that is restricted" is a complete question on its own,
  // so it runs without a term where every other filter would need one.
  if (view.request.term === '' && view.request.restrictedOnly !== true) {
    handle.setBusy(view.strings.prompt);
  }

  return handle;
}

function buildMode(
  doc: Document,
  host: HTMLElement,
  label: string,
  checked: boolean
): HTMLInputElement {
  const wrapper: HTMLElement = doc.createElement('label');
  wrapper.className = 'mdv-tag-option';

  const input: HTMLInputElement = doc.createElement('input');
  input.setAttribute('type', 'radio');
  input.setAttribute('name', 'mdv-search-mode');
  input.checked = checked;
  wrapper.appendChild(input);

  const text: HTMLElement = doc.createElement('span');
  text.textContent = label;
  wrapper.appendChild(text);

  host.appendChild(wrapper);
  return input;
}

/**
 * The hits, as the tree they came out of.
 *
 * A flat list loses which section a hit belongs to, and in a wiki that is often
 * the thing that tells you whether it is the one you wanted. A body search adds
 * the matching line under each file.
 */
function renderResults(
  doc: Document,
  host: HTMLElement,
  hits: ISearchHit[],
  view: ISearchView
): void {
  host.innerHTML = '';

  const count: HTMLElement = doc.createElement('p');
  count.className = 'mdv-search-count';
  count.setAttribute('role', 'status');
  count.textContent = view.strings.resultCount.replace('{0}', String(hits.length));
  host.appendChild(count);

  if (hits.length === 0) {
    const empty: HTMLElement = doc.createElement('p');
    empty.className = 'mdv-related-empty';
    empty.textContent = view.strings.empty;
    host.appendChild(empty);
    return;
  }

  const pages: IRelatedPage[] = [];
  for (let i: number = 0; i < hits.length; i++) {
    pages.push(hits[i].page);
  }

  // Rooted at the deepest folder they share, and named — otherwise the rows
  // hang off nothing and the reader cannot tell which section they are in.
  const root: string = commonRoot(pages);
  const tree: ITreeFolder = buildTree(pages, root);

  if (!isEmptyTree(tree)) {
    host.appendChild(
      renderTree(doc, tree, {
        buildDocUrl: view.buildDocUrl,
        linkAttribute: view.linkAttribute,
        currentPath: '',
        rootLabel: root === '' ? undefined : root + '/',
        linkFolders: true,
        lockedLabel: view.strings.locked
      })
    );
  }

  // Snippets go underneath rather than inside the tree: a wrapped line of prose
  // in a box-drawn tree destroys the alignment the rules exist for.
  const withSnippets: ISearchHit[] = [];
  for (let i: number = 0; i < hits.length; i++) {
    if (hits[i].snippet !== '') {
      withSnippets.push(hits[i]);
    }
  }

  if (withSnippets.length === 0) {
    return;
  }

  const list: HTMLElement = doc.createElement('div');
  list.className = 'mdv-search-snippets';

  for (let i: number = 0; i < withSnippets.length; i++) {
    list.appendChild(buildSnippetRow(doc, withSnippets[i], view));
  }

  host.appendChild(list);
}

function buildSnippetRow(doc: Document, hit: ISearchHit, view: ISearchView): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-search-snippet';

  const link: HTMLElement = doc.createElement('a');
  link.className = 'mdv-related-link';
  link.setAttribute('href', view.buildDocUrl(hit.page.path));
  link.setAttribute(view.linkAttribute, hit.page.path);
  link.textContent = hit.page.path;
  row.appendChild(link);

  const text: HTMLElement = doc.createElement('div');
  text.className = 'mdv-search-snippet-text';
  text.textContent = hit.snippet;
  row.appendChild(text);

  return row;
}

/**
 * The deepest folder every hit sits under, so the tree is not padded out with
 * levels that carry a single child all the way down.
 */
export function commonRoot(pages: IRelatedPage[]): string {
  if (pages.length === 0) {
    return '';
  }

  let common: string[] = dirname(pages[0].path).split('/');
  if (common.length === 1 && common[0] === '') {
    return '';
  }

  for (let i: number = 1; i < pages.length; i++) {
    const parts: string[] = dirname(pages[i].path).split('/');
    const shared: string[] = [];

    for (let s: number = 0; s < common.length && s < parts.length; s++) {
      if (common[s].toLowerCase() !== parts[s].toLowerCase()) {
        break;
      }
      shared.push(common[s]);
    }

    common = shared;
    if (common.length === 0) {
      return '';
    }
  }

  return common.join('/');
}
