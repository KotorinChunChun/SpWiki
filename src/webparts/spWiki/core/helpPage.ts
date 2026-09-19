/**
 * The help page.
 *
 * Two things live here, and they are different in kind:
 *
 *  - the documentation, baked into the bundle so it is available wherever the
 *    web part is, with nothing for an administrator to deploy and nothing to
 *    go stale against the version actually running;
 *  - the values needed to register the list view command set, built from the
 *    page the reader is on. Pointing `viewerPageUrl` at the wrong site cost a
 *    real afternoon — the commands appear and simply do nothing — so the JSON
 *    is produced rather than described.
 *
 * It is a page rather than a panel: it replaces the document instead of pushing
 * it down, and it has its own `?help=` address, so a section of the manual can
 * be linked to like anything else.
 */

import { IBundledDoc } from './bundledDocs';

/** Component id of the list view command set, from its manifest. */
export const COMMAND_SET_COMPONENT_ID: string = '7bf4c5c5-631b-48cf-b066-e92bfb4c91dc';

/** Location value that puts the commands on both the toolbar and the menu. */
export const COMMAND_SET_LOCATION: string = 'ClientSideExtension.ListViewCommandSet';

/** List template of a document library. */
export const COMMAND_SET_LIST_TEMPLATE: string = '101';

export interface IHelpSetup {
  /** Server relative URL of the page the web part sits on. */
  viewerPageUrl: string;
  /** Server relative URL of the library, when it is not the obvious one. */
  libraryRootUrl: string;
  /** `env` value the page is being viewed with. */
  chromeMode: string;
  /** Address of the tenant wide extensions list, or '' when not known. */
  extensionsListUrl: string;
}

export interface IHelpStrings {
  title: string;
  /** Heading above the list of documents. */
  documents: string;
  /** Heading of the command set section. */
  setupTitle: string;
  /** Sentence above the command set values. */
  setupIntro: string;
  componentId: string;
  location: string;
  listTemplate: string;
  componentProperties: string;
  extensionsList: string;
  copy: string;
  copied: string;
  /** Link back from a document to the help index. */
  backToIndex: string;
}

export interface IHelpPageView {
  container: HTMLElement;
  /** Which document to show; '' shows the index. */
  docKey: string;
  docs: IBundledDoc[];
  setup: IHelpSetup;
  strings: IHelpStrings;
  /** Address of `?help=<key>`, so the entries are real links. */
  buildHelpUrl(key: string): string;
  /** Renders one document's Markdown into the given element. */
  renderMarkdown(markdown: string, into: HTMLElement): void;
  /** Puts a value on the clipboard; the page handles the feedback. */
  onCopy(value: string): void;
}

/**
 * The component properties JSON for this page, ready to paste.
 *
 * `libraryRootUrl` is only written when the library is not the one the command
 * would default to anyway, so the common case stays a two line object rather
 * than something the reader has to check.
 */
export function buildComponentProperties(setup: IHelpSetup): string {
  const lines: string[] = ['  "viewerPageUrl": ' + JSON.stringify(setup.viewerPageUrl)];

  if (setup.libraryRootUrl !== '') {
    lines.push('  "libraryRootUrl": ' + JSON.stringify(setup.libraryRootUrl));
  }

  lines.push('  "chromeMode": ' + JSON.stringify(setup.chromeMode || 'WebView'));

  return '{\n' + lines.join(',\n') + '\n}';
}

export function renderHelpPage(view: IHelpPageView): void {
  const doc: Document = view.container.ownerDocument;
  view.container.innerHTML = '';

  const selected: IBundledDoc | undefined = findDoc(view.docs, view.docKey);

  if (selected) {
    renderDocument(doc, view, selected);
    return;
  }

  renderIndex(doc, view);
}

function findDoc(docs: IBundledDoc[], key: string): IBundledDoc | undefined {
  for (let i: number = 0; i < docs.length; i++) {
    if (docs[i].key === key) {
      return docs[i];
    }
  }
  return undefined;
}

/** One document, with a way back to the list it came from. */
function renderDocument(doc: Document, view: IHelpPageView, selected: IBundledDoc): void {
  const back: HTMLElement = doc.createElement('a');
  back.className = 'mdv-help-back';
  back.setAttribute('href', view.buildHelpUrl(''));
  back.setAttribute('data-mdv-help', '');
  back.textContent = view.strings.backToIndex;
  view.container.appendChild(back);

  const body: HTMLElement = doc.createElement('div');
  body.className = 'mdv-help-doc';
  view.container.appendChild(body);

  view.renderMarkdown(selected.markdown, body);
}

/** The list of documents, then the values needed to set the commands up. */
function renderIndex(doc: Document, view: IHelpPageView): void {
  const heading: HTMLElement = doc.createElement('h1');
  heading.textContent = view.strings.title;
  view.container.appendChild(heading);

  const docsHeading: HTMLElement = doc.createElement('h2');
  docsHeading.textContent = view.strings.documents;
  view.container.appendChild(docsHeading);

  const list: HTMLElement = doc.createElement('ul');
  list.className = 'mdv-help-docs';

  for (let i: number = 0; i < view.docs.length; i++) {
    const item: HTMLElement = doc.createElement('li');

    const link: HTMLElement = doc.createElement('a');
    link.className = 'mdv-related-link';
    link.setAttribute('href', view.buildHelpUrl(view.docs[i].key));
    link.setAttribute('data-mdv-help', view.docs[i].key);
    link.textContent = view.docs[i].title;
    item.appendChild(link);

    list.appendChild(item);
  }

  view.container.appendChild(list);

  const setupHeading: HTMLElement = doc.createElement('h2');
  setupHeading.textContent = view.strings.setupTitle;
  view.container.appendChild(setupHeading);

  const intro: HTMLElement = doc.createElement('p');
  intro.className = 'mdv-help-intro';
  intro.textContent = view.strings.setupIntro;
  view.container.appendChild(intro);

  view.container.appendChild(buildRow(doc, view, view.strings.componentId, COMMAND_SET_COMPONENT_ID, false));
  view.container.appendChild(buildRow(doc, view, view.strings.location, COMMAND_SET_LOCATION, false));
  view.container.appendChild(
    buildRow(doc, view, view.strings.listTemplate, COMMAND_SET_LIST_TEMPLATE, false)
  );
  view.container.appendChild(
    buildRow(doc, view, view.strings.componentProperties, buildComponentProperties(view.setup), true)
  );

  if (view.setup.extensionsListUrl !== '') {
    view.container.appendChild(
      buildRow(doc, view, view.strings.extensionsList, view.setup.extensionsListUrl, false)
    );
  }
}

/** A labelled value with its own copy button. */
function buildRow(
  doc: Document,
  view: IHelpPageView,
  label: string,
  value: string,
  multiline: boolean
): HTMLElement {
  const row: HTMLElement = doc.createElement('div');
  row.className = 'mdv-help-row';

  const name: HTMLElement = doc.createElement('div');
  name.className = 'mdv-help-label';
  name.textContent = label;
  row.appendChild(name);

  const body: HTMLElement = doc.createElement('div');
  body.className = 'mdv-help-value-row';

  const box: HTMLElement = doc.createElement(multiline ? 'pre' : 'code');
  box.className = 'mdv-help-value';
  box.textContent = value;
  body.appendChild(box);

  const copy: HTMLButtonElement = doc.createElement('button');
  copy.className = 'mdv-btn mdv-help-copy';
  copy.setAttribute('type', 'button');
  copy.textContent = view.strings.copy;
  copy.onclick = (): void => {
    view.onCopy(value);

    // Confirming on the button itself: the value is right there, and a toast
    // for something this small would be more movement than news.
    copy.textContent = view.strings.copied;
    const win: Window | null = doc.defaultView;
    if (win) {
      win.setTimeout((): void => {
        copy.textContent = view.strings.copy;
      }, 1200);
    }
  };
  body.appendChild(copy);

  row.appendChild(body);
  return row;
}

/**
 * The `?help=` key of a help link that was clicked, or undefined when the click
 * was somewhere else. An empty string means the index.
 */
// `null` here mirrors the DOM APIs this walks (`event.target`, `parentNode`).
// eslint-disable-next-line @rushstack/no-new-null
export function findHelpLink(start: Node | null, root: Node): string | undefined {
  let node: Node | null = start;

  while (node && node !== root) {
    if (node.nodeType === 1 && (node as Element).hasAttribute('data-mdv-help')) {
      return (node as Element).getAttribute('data-mdv-help') || '';
    }
    node = node.parentNode;
  }

  return undefined;
}
