/**
 * Local preview of SpWiki (which also renders `.html` files).
 *
 * It reuses the same core modules as the web part (rendering, sanitizing,
 * relative link resolution, toolbar, tag bar and search, headings, navigation
 * popup, related pages, tag editor) but reads the documents from the
 * `tests/fixtures/documents` folder over plain `fetch`, and keeps the library listing, the
 * tags and any file it creates or renames in memory instead of in SharePoint.
 * That makes it possible to check everything except the REST calls without a
 * tenant.
 */
import { renderDocument } from '../../src/webparts/spWiki/core/renderer';
import { createSanitizer } from '../../src/webparts/spWiki/core/sanitize';
import { findDocLinkPath, isPlainLeftClick } from '../../src/webparts/spWiki/core/domRewrite';
import { findHeadingAnchor, toggleHeadingAt } from '../../src/webparts/spWiki/core/headings';
import { findCodeToCopy } from '../../src/webparts/spWiki/core/codeBlocks';
import { buildHeadingNav } from '../../src/webparts/spWiki/core/toc';
import { findDirectives, renderDirectives } from '../../src/webparts/spWiki/core/directives';
import { BUNDLED_DOCS } from '../../src/webparts/spWiki/core/bundledDocs';
import { renderHelpPage } from '../../src/webparts/spWiki/core/helpPage';
import {
  dirname,
  encodePath,
  isDocumentPath,
  joinUrl,
  normalizePath,
  sanitizeDocPath,
  splitLink
} from '../../src/webparts/spWiki/core/pathUtils';
import { getQueryParam, removeQueryParam, setQueryParam } from '../../src/webparts/spWiki/core/urlUtils';
import {
  DOC_LINK_ATTRIBUTE,
  IDocLinkContext,
  IDocumentResult,
  Sanitizer
} from '../../src/webparts/spWiki/core/types';
import {
  ITagFieldInfo,
  ITagGroup,
  flattenTagGroups,
  mergeTagGroups
} from '../../src/webparts/spWiki/core/tags';
import {
  IRelatedPage,
  filterPagesByTag,
  renderTagSearch
} from '../../src/webparts/spWiki/core/relatedPages';
import {
  INavEntry,
  INewPageRequest,
  buildAncestors,
  renderNavPanel,
  renderRenamePanel,
  withExtension
} from '../../src/webparts/spWiki/core/navPanel';
import { newDocumentTemplate } from '../../src/webparts/spWiki/core/templates';
import {
  ACTION_ATTRIBUTE,
  IToolbarButton,
  ToolbarAction,
  renderNotice,
  showToast,
  renderTagBar,
  renderTagEditor,
  renderToolbar
} from '../../src/webparts/spWiki/core/chrome';
import { buildProtocolUrl } from '../../src/webparts/spWiki/core/appLink';

interface IMockFolder {
  folders: string[];
  files: string[];
}

interface IMockPage {
  path: string;
  name: string;
  tags: { [internalName: string]: string[] };
}

interface IMockMetadata {
  fields: ITagFieldInfo[];
  folders: { [path: string]: IMockFolder };
  pages: IMockPage[];
}

const QUERY_PARAMETER: string = 'file';
const TAG_PARAMETER: string = 'tag';
const LIBRARY_ROOT: string = '/tests/fixtures/documents';
const START_FILE: string = 'README.md';
const TITLE: string = 'ドキュメント';

/** Stand-ins for the values the web part reads from the SharePoint context. */
const MOCK_SITE_URL: string = 'https://tenant.sharepoint.com/sites/docs';
const MOCK_LIBRARY_URL: string = '/sites/docs/Shared Documents';

const sanitize: Sanitizer = createSanitizer(window);

const root: HTMLElement = document.getElementById('viewer') as HTMLElement;
const toolbar: HTMLElement = document.getElementById('toolbar') as HTMLElement;
const tagBar: HTMLElement = document.getElementById('tagBar') as HTMLElement;
const navHost: HTMLElement = document.getElementById('nav-panel') as HTMLElement;
const notice: HTMLElement = document.getElementById('notice') as HTMLElement;
const tagPanel: HTMLElement = document.getElementById('tagPanel') as HTMLElement;
const renamePanel: HTMLElement = document.getElementById('renamePanel') as HTMLElement;
const toasts: HTMLElement = document.getElementById('toasts') as HTMLElement;
const content: HTMLElement = document.getElementById('content') as HTMLElement;
const body: HTMLElement = document.getElementById('body') as HTMLElement;
const outlinePopup: HTMLElement = document.getElementById('outlinePopup') as HTMLElement;
const sidePane: HTMLElement = document.getElementById('sidePane') as HTMLElement;
const topPanel: HTMLElement = document.getElementById('topPanel') as HTMLElement;
const footerPath: HTMLElement = document.getElementById('footerPath') as HTMLElement;

let metadata: IMockMetadata = { fields: [], folders: {}, pages: [] };
let currentPath: string = '';
let loadToken: number = 0;
let tagPanelOpen: boolean = false;
/** Internal name of the column the open tag editor is for. */
let tagPanelField: string = '';
let navPanelOpen: boolean = false;
let renamePanelOpen: boolean = false;
let helpOpen: boolean = false;
type SidePanel = 'outline' | 'tree';
let pinnedPanel: SidePanel | undefined;
let outlinePopupOpen: boolean = false;
let outlineClickTimer: number = 0;
let menuClickTimer: number = 0;

/** Documents created through the popup; the harness has nowhere to write them. */
const createdFiles: { [path: string]: string } = {};

// ------------------------------------------------------------------ loading

async function fetchDocument(path: string): Promise<IDocumentResult> {
  if (Object.prototype.hasOwnProperty.call(createdFiles, path)) {
    return { kind: 'ok', text: createdFiles[path] };
  }

  try {
    const response: Response = await fetch(joinUrl(LIBRARY_ROOT, encodePath(path)));
    if (response.status === 404) {
      return { kind: 'notfound' };
    }
    if (!response.ok) {
      return { kind: 'error', message: 'HTTP ' + response.status };
    }
    return { kind: 'ok', text: await response.text() };
  } catch (error) {
    return { kind: 'error', message: String(error) };
  }
}

async function loadMetadata(): Promise<void> {
  const response: Response = await fetch('/tests/harness/mock-metadata.json');
  metadata = await response.json();
}

function findPage(path: string): IMockPage | undefined {
  for (let i: number = 0; i < metadata.pages.length; i++) {
    if (metadata.pages[i].path.toLowerCase() === path.toLowerCase()) {
      return metadata.pages[i];
    }
  }
  return undefined;
}

/** The tags of one document, one group per configured column. */
function getTagGroups(path: string): ITagGroup[] {
  const page: IMockPage | undefined = findPage(path);
  const groups: ITagGroup[] = [];

  for (let i: number = 0; i < metadata.fields.length; i++) {
    const field: ITagFieldInfo = metadata.fields[i];
    const tags: string[] = page && page.tags[field.internalName] ? page.tags[field.internalName] : [];
    groups.push({ field: field, tags: tags });
  }

  return groups;
}

function setTagGroups(path: string, groups: ITagGroup[]): void {
  let page: IMockPage | undefined = findPage(path);

  if (!page) {
    page = { path: path, name: path.split('/').pop() || path, tags: {} };
    metadata.pages.push(page);
  }

  // The column's own choices are left alone, exactly as SharePoint leaves them:
  // a filled-in value lives on the item, not in the column definition. The
  // editor shows it anyway because it merges the item's tags into the list.
  for (let i: number = 0; i < groups.length; i++) {
    page.tags[groups[i].field.internalName] = groups[i].tags;
  }
}

/** Every document in the library with the union of its tags. */
function getPages(): IRelatedPage[] {
  const pages: IRelatedPage[] = [];

  for (let i: number = 0; i < metadata.pages.length; i++) {
    const page: IMockPage = metadata.pages[i];
    let tags: string[] = [];

    for (let f: number = 0; f < metadata.fields.length; f++) {
      const value: string[] = page.tags[metadata.fields[f].internalName] || [];
      tags = tags.concat(value);
    }

    pages.push({ path: page.path, name: page.name, tags: tags });
  }

  return pages;
}

function listFolder(folder: string): { folders: INavEntry[]; files: INavEntry[] } {
  const entry: IMockFolder = metadata.folders[folder] || { folders: [], files: [] };
  const folders: INavEntry[] = [];
  const files: INavEntry[] = [];

  for (let i: number = 0; i < entry.folders.length; i++) {
    const name: string = entry.folders[i];
    folders.push({ name: name, path: folder === '' ? name : folder + '/' + name });
  }

  for (let i: number = 0; i < entry.files.length; i++) {
    const name: string = entry.files[i];
    if (isDocumentPath(name)) {
      files.push({ name: name, path: folder === '' ? name : folder + '/' + name });
    }
  }

  return { folders: folders, files: files };
}

// ----------------------------------------------------------------- rendering

/**
 * Built from scratch rather than edited, exactly as the web part does it: only
 * the chrome setting survives a navigation, so a finished search or help
 * request cannot ride along into the next document's address.
 */
function buildDocUrl(path: string): string {
  const search: string = setQueryParam('', QUERY_PARAMETER, path);
  const chrome: string = getQueryParam(window.location.search, 'env');

  return (
    window.location.pathname +
    (chrome === '' ? search : setQueryParam(search, 'env', chrome))
  );
}

function buildTagUrl(tag: string): string {
  return (
    window.location.pathname +
    setQueryParam(removeQueryParam(window.location.search, QUERY_PARAMETER), TAG_PARAMETER, tag)
  );
}

function createContext(path: string): IDocLinkContext {
  return {
    currentPath: path,
    buildDocUrl: buildDocUrl,
    buildFileUrl: (filePath: string): string => joinUrl(LIBRARY_ROOT, encodePath(filePath))
  };
}

function showStatus(className: string, message: string): void {
  content.innerHTML = '';
  const div: HTMLDivElement = document.createElement('div');
  div.className = className;
  div.textContent = message;
  content.appendChild(div);
}

function drawToolbar(ready: boolean): void {
  const buttons: IToolbarButton[] = [
    { action: 'edit', label: '編集', disabled: !ready },
    { action: 'open-app', label: 'アプリで開く', disabled: !ready }
  ];

  renderToolbar({
    container: toolbar,
    path: currentPath,
    buttons: buttons,
    menuLabel: 'メニュー',
    title: TITLE,
    titleLabel: '開始ページに戻る',
    rootLabel: '(ライブラリのルート)',
    pathLabel: ready ? 'クリックしてファイル名・フォルダーを変更' : undefined,
    outline: { label: '目次', pinned: pinnedPanel === 'outline' },
    menuPinned: pinnedPanel === 'tree',
    folderLinkAttribute: DOC_LINK_ATTRIBUTE,
    buildFolderUrl: buildDocUrl
  });
}

function drawTagBar(): void {
  renderTagBar({
    container: tagBar,
    groups: getTagGroups(currentPath),
    buildTagUrl: buildTagUrl,
    onTag: (tag: string): void => searchTag(tag, true),
    onEdit: (internalName: string): void => {
      if (tagPanelOpen && tagPanelField === internalName) {
        closePanels();
      } else {
        openTagPanel(internalName);
      }
    },
    editLabel: 'クリックしてタグを編集'
  });
}

async function navigate(path: string, pushHistory: boolean, hash?: string): Promise<void> {
  closePanels();
  notice.innerHTML = '';
  tagBar.innerHTML = '';

  if (path === '') {
    showStatus('status', 'ファイルが指定されていません。');
    return;
  }

  currentPath = path;
  footerPath.textContent = path;
  updateActiveNavigation(path);
  drawToolbar(false);

  if (pushHistory) {
    window.history.pushState({ file: path }, '', buildDocUrl(path) + (hash || ''));
  }

  loadToken += 1;
  const token: number = loadToken;
  showStatus('status', 'Loading...');

  const result: IDocumentResult = await fetchDocument(path);
  if (token !== loadToken) {
    return;
  }

  if (result.kind === 'notfound') {
    showStatus('status', 'ファイルが見つかりません。');
    drawToolbar(false);
    return;
  }

  if (result.kind === 'error') {
    showStatus('error', 'ファイルを読み込めませんでした。 ' + (result.message || ''));
    return;
  }

  renderDocument({
    source: result.text || '',
    container: content,
    sanitize: sanitize,
    context: createContext(path),
    headings: {
      copyLabel: 'この見出しへのリンクをコピー',
      toggleLabel: 'クリックで折りたたみ'
    },
    codeBlocks: { copyLabel: 'コードをコピー' }
  });

  renderDirectives(content, findDirectives(content), {
    pages: getPages(),
    currentPath: path,
    buildDocUrl: buildDocUrl,
    linkAttribute: DOC_LINK_ATTRIBUTE,
    strings: { empty: '該当するページはありません。', rootLabel: '(ライブラリのルート)' }
  });

  drawToolbar(true);
  drawTagBar();
  renderSidePanels();
  scrollToAnchor(hash !== undefined ? hash : window.location.hash);
}

function searchTag(tag: string, pushHistory: boolean): void {
  closePanels();
  notice.innerHTML = '';
  tagBar.innerHTML = '';

  if (pushHistory) {
    window.history.pushState({ tag: tag }, '', buildTagUrl(tag));
  }

  loadToken += 1;

  renderTagSearch({
    container: content,
    tag: tag,
    pages: filterPagesByTag(getPages(), tag),
    title: 'タグ:',
    emptyMessage: 'このタグが付いたページはありません。',
    linkAttribute: DOC_LINK_ATTRIBUTE,
    buildDocUrl: buildDocUrl
  });
}

/** Brings the heading a `#...` in the address refers to into view. */
function scrollToAnchor(hash: string): void {
  if (!hash) {
    return;
  }

  let id: string = hash.charAt(0) === '#' ? hash.slice(1) : hash;
  try {
    id = decodeURIComponent(id);
  } catch {
    // Malformed percent-encoding: match on what was given.
  }

  const candidates: NodeListOf<Element> = content.querySelectorAll('[id]');
  for (let i: number = 0; i < candidates.length; i++) {
    if (candidates[i].getAttribute('id') === id) {
      candidates[i].scrollIntoView();
      return;
    }
  }
}

function updateActiveNavigation(path: string): void {
  const links: NodeListOf<Element> = document.querySelectorAll('#nav a[data-path]');
  for (let i: number = 0; i < links.length; i++) {
    const link: Element = links[i];
    if (link.getAttribute('data-path') === path) {
      link.setAttribute('class', 'active');
    } else {
      link.removeAttribute('class');
    }
  }
}

function readPathFromUrl(): string {
  const fromQuery: string = sanitizeDocPath(getQueryParam(window.location.search, QUERY_PARAMETER));
  return fromQuery !== '' ? fromQuery : START_FILE;
}

// ------------------------------------------------------------------- actions

function showNotice(message: string, detail?: string, copyValue?: string): void {
  renderNotice({
    container: notice,
    message: message,
    detail: detail,
    copyLabel: copyValue ? 'パスをコピー' : undefined,
    onCopy: copyValue
      ? (): void => {
          void navigator.clipboard.writeText(copyValue);
        }
      : undefined,
    closeLabel: '閉じる',
    onClose: (): void => {
      notice.innerHTML = '';
    }
  });
}

function closePanels(): void {
  tagPanelOpen = false;
  tagPanelField = '';
  renamePanelOpen = false;
  navPanelOpen = false;
  tagPanel.innerHTML = '';
  renamePanel.innerHTML = '';
  helpOpen = false;
  navHost.innerHTML = '';
}

function openTagPanel(internalName: string): void {
  closePanels();
  const path: string = currentPath;

  let group: ITagGroup | undefined;
  const groups: ITagGroup[] = getTagGroups(path);
  for (let i: number = 0; i < groups.length; i++) {
    if (groups[i].field.internalName === internalName) {
      group = groups[i];
      break;
    }
  }

  if (!group) {
    showNotice('タグの列が見つかりません。');
    return;
  }

  tagPanelOpen = true;
  tagPanelField = internalName;

  renderTagEditor({
    container: tagPanel,
    groups: [group],
    strings: {
      title: 'タグを編集',
      save: '保存',
      cancel: 'キャンセル',
      textHint: '複数のタグは ; か , で区切ってください。',
      noChoices: 'この列に選択肢が登録されていません。',
      addLabel: '新しいタグを追加',
      addPlaceholder: '例: 設計; 手順書',
      addDisabledPlaceholder: 'この列にはタグを追加できません',
      addHint: '一覧に無いタグはここに入力します。複数は ; か , で区切ってください。',
      addDisabledHint: 'この列は選択肢以外の値を受け付けません。列の設定で「選択肢に含まれていない値の入力を許可する」を有効にすると、ここから追加できるようになります。'
    },
    onSave: (edited: ITagGroup[]): void => {
      // Only the open column comes back, exactly as in the web part.
      setTagGroups(path, mergeTagGroups(getTagGroups(path), edited));
      closePanels();
      drawTagBar();
    },
    onCancel: closePanels
  });
}

function openNavPanel(folder: string): void {
  const listing: { folders: INavEntry[]; files: INavEntry[] } = listFolder(folder);

  // The same panel in two hosts: the popup under the button, or the pinned pane.
  const host: HTMLElement = pinnedPanel === 'tree' ? sidePane : navHost;
  navPanelOpen = pinnedPanel !== 'tree';

  renderNavPanel({
    container: host,
    ancestors: buildAncestors(folder),
    folders: listing.folders,
    files: listing.files,
    currentPath: currentPath,
    strings: {
      newPage: '新しいページを作成',
      newPageName: 'ファイル名',
      format: '形式',
      openWith: '作成後',
      openInSharePoint: 'SharePoint で編集',
      openInApp: 'アプリで編集',
      create: '作成',
      cancel: 'キャンセル',
      location: '現在の場所',
      root: '(ライブラリのルート)',
      empty: 'このフォルダーに表示できるファイルはありません。',
      search: '検索',
      help: 'ヘルプ',
      chromeToggle: '全画面表示に切り替える',
      close: '閉じる'
    },
    buildDocUrl: buildDocUrl,
    onOpen: (path: string): void => {
      closePanels();
      void navigate(path, true);
    },
    onFolder: (path: string): void => openNavPanel(path),
    onCreate: (request: INewPageRequest): void => createPage(folder, request),
    onClose: closePanels,
    onSearch: (): void => {
      closePanels();
      showToast({ container: toasts, message: 'ハーネスでは検索ページは動きません' });
    },
    onHelp: (): void => {
      closePanels();
      openHelp('');
    },
    // The harness has no SharePoint chrome to switch off.
    onChromeToggle: (): void => {
      showToast({ container: toasts, message: 'ハーネスでは表示を切り替えられません' });
    }
  });
}

/** The help page, in the content area exactly as in the web part. */
function openHelp(docKey: string): void {
  closePanels();
  helpOpen = true;
  tagBar.innerHTML = '';
  content.innerHTML = '';
  footerPath.textContent = '';

  renderHelpPage({
    container: content,
    docKey: docKey,
    docs: BUNDLED_DOCS,
    setup: {
      viewerPageUrl: '/sites/docs/SitePages/SpWiki.aspx',
      libraryRootUrl: '',
      chromeMode: 'WebView',
      extensionsListUrl: 'https://tenant.sharepoint.com/sites/appcatalog/Lists/TenantWideExtensions'
    },
    strings: {
      title: 'ヘルプ',
      documents: 'ドキュメント',
      setupTitle: 'ライブラリのコマンドの設定',
      setupIntro:
        'ライブラリのツールバーと右クリック メニューにコマンドを出すには、アプリ カタログの「テナント全体の拡張機能」リストに次の値で 1 行追加します。',
      componentId: 'コンポーネント ID',
      location: '場所',
      listTemplate: 'リストのテンプレート',
      componentProperties: 'コンポーネントのプロパティ',
      extensionsList: 'テナント全体の拡張機能リスト',
      copy: 'コピー',
      copied: 'コピーしました',
      backToIndex: '← ヘルプの目次へ'
    },
    buildHelpUrl: (key: string): string => '?help=' + key,
    renderMarkdown: (markdown: string, into: HTMLElement): void => {
      renderDocument({
        source: markdown,
        container: into,
        sanitize: sanitize,
        format: 'markdown',
        context: {
          currentPath: '',
          buildDocUrl: (): string => '',
          buildFileUrl: (): string => ''
        },
        headings: { copyLabel: 'この見出しへのリンクをコピー', toggleLabel: 'クリックで折りたたみ' },
        codeBlocks: { copyLabel: 'コードをコピー' }
      });
    },
    onCopy: (value: string): void => {
      void navigator.clipboard.writeText(value).catch((): void => undefined);
    }
  });
}

function createPage(folder: string, request: INewPageRequest): void {
  const fileName: string = withExtension(request.name, request.format);
  if (fileName === '') {
    showNotice('拡張子は .md / .markdown / .html / .htm のいずれかにしてください。');
    return;
  }

  const path: string = folder === '' ? fileName : folder + '/' + fileName;
  if (!metadata.folders[folder]) {
    metadata.folders[folder] = { folders: [], files: [] };
  }
  if (metadata.folders[folder].files.indexOf(fileName) < 0) {
    metadata.folders[folder].files.push(fileName);
  }
  createdFiles[path] = newDocumentTemplate(fileName, request.format);

  closePanels();
  void navigate(path, true).then((): void => {
    showNotice(
      request.editor === 'app'
        ? 'テンプレートから作成しました（ハーネスではメモリ上のみ）。実際にはデスクトップ アプリが起動します。'
        : 'テンプレートから作成しました（ハーネスではメモリ上のみ）。実際には SharePoint の編集画面へ遷移します。'
    );
  });
}

function openRenamePanel(): void {
  closePanels();
  renamePanelOpen = true;

  renderRenamePanel({
    container: renamePanel,
    path: currentPath,
    folders: Object.keys(metadata.folders).filter((f: string): boolean => f !== ''),
    strings: {
      title: 'ファイル名 / フォルダーの変更',
      folder: 'フォルダー',
      name: 'ファイル名',
      hint: 'ライブラリのルートからの相対パスです。移動先のフォルダーは既に存在している必要があります。',
      save: '保存',
      cancel: 'キャンセル'
    },
    onSave: (path: string): void => renameTo(path),
    onCancel: closePanels
  });
}

function renameTo(rawPath: string): void {
  const path: string = normalizePath(rawPath);

  if (path === '' || !isDocumentPath(path)) {
    showNotice('拡張子は .md / .markdown / .html / .htm のいずれかにしてください。');
    return;
  }

  if (path.toLowerCase() === currentPath.toLowerCase()) {
    closePanels();
    return;
  }

  // The harness only moves the entry, never the file behind it, so the content
  // of a renamed sample document is served from the copy kept in memory.
  const source: string = currentPath;
  const page: IMockPage | undefined = findPage(source);
  if (page) {
    page.path = path;
    page.name = path.split('/').pop() || path;
  }

  const oldFolder: string = dirname(source);
  const oldName: string = source.split('/').pop() || source;
  if (metadata.folders[oldFolder]) {
    const index: number = metadata.folders[oldFolder].files.indexOf(oldName);
    if (index >= 0) {
      metadata.folders[oldFolder].files.splice(index, 1);
    }
  }

  const folder: string = dirname(path);
  const name: string = path.split('/').pop() || path;
  if (!metadata.folders[folder]) {
    metadata.folders[folder] = { folders: [], files: [] };
  }
  metadata.folders[folder].files.push(name);

  closePanels();

  void fetchDocument(source).then((result: IDocumentResult): void => {
    createdFiles[path] = result.text || '';
    void navigate(path, true).then((): void => {
      showToast({ container: toasts, message: 'ファイル名を変更しました' });
    });
  });
}

function openInApp(): void {
  const fileUrl: string = MOCK_SITE_URL + '/Shared%20Documents/' + encodePath(currentPath);

  showNotice('ハーネスではスキームを起動しません。実際には ' + buildProtocolUrl('spmd', fileUrl) + ' を開きます。');
}

/** The same deferred single click / double click split as the web part. */
function onOutlineClick(): void {
  if (outlineClickTimer) {
    window.clearTimeout(outlineClickTimer);
  }
  outlineClickTimer = window.setTimeout((): void => {
    outlineClickTimer = 0;
    if (pinnedPanel === 'outline') {
      pinPanel(undefined);
      return;
    }
    outlinePopupOpen = !outlinePopupOpen;
    renderSidePanels();
  }, 250);
}

function onMenuClick(): void {
  if (menuClickTimer) {
    window.clearTimeout(menuClickTimer);
  }
  menuClickTimer = window.setTimeout((): void => {
    menuClickTimer = 0;
    if (pinnedPanel === 'tree') {
      pinPanel(undefined);
      return;
    }
    if (navPanelOpen) {
      closePanels();
    } else {
      openNavPanel(dirname(currentPath));
    }
  }, 250);
}

/** One pane, one occupant: pinning the second replaces the first. */
function pinPanel(panel: SidePanel | undefined): void {
  pinnedPanel = panel;
  outlinePopupOpen = false;

  if (panel !== 'tree') {
    navPanelOpen = false;
    navHost.innerHTML = '';
  }

  renderSidePanels();
  drawToolbar(true);
}

function renderSidePanels(): void {
  body.className = pinnedPanel ? 'body body-pinned' : 'body';

  if (pinnedPanel !== 'tree') {
    sidePane.innerHTML = '';
  }

  renderOutline();

  if (pinnedPanel === 'tree') {
    openNavPanel(dirname(currentPath));
  }

  measureTopPanel();
}

/** The side pane starts where the sticky chrome ends, so that is measured. */
function measureTopPanel(): void {
  root.style.setProperty('--mdv-top', topPanel.offsetHeight + 'px');
}

function renderOutline(): void {
  outlinePopup.innerHTML = '';

  const pinned: boolean = pinnedPanel === 'outline';
  if (!pinned && !outlinePopupOpen) {
    return;
  }

  const host: HTMLElement = pinned ? sidePane : outlinePopup;
  if (pinned) {
    host.innerHTML = '';
  }
  const panel: HTMLElement = document.createElement('div');
  panel.className = 'mdv-outline-panel';

  const heading: HTMLElement = document.createElement('div');
  heading.className = 'mdv-outline-title';
  heading.textContent = '目次';
  panel.appendChild(heading);

  const nav: HTMLElement | undefined = buildHeadingNav(content);
  if (nav) {
    panel.appendChild(nav);
  } else {
    const empty: HTMLElement = document.createElement('p');
    empty.className = 'mdv-related-empty';
    empty.textContent = 'このページには見出しがありません。';
    panel.appendChild(empty);
  }

  host.appendChild(panel);
}

function runAction(action: ToolbarAction): void {
  if (action === 'outline') {
    onOutlineClick();
    return;
  }

  if (action === 'home') {
    void navigate(START_FILE, true);
    return;
  }

  if (action === 'menu') {
    if (navPanelOpen) {
      closePanels();
    } else {
      openNavPanel(dirname(currentPath));
    }
    return;
  }

  if (action === 'path') {
    if (renamePanelOpen) {
      closePanels();
    } else {
      openRenamePanel();
    }
    return;
  }

  if (action === 'edit') {
    showNotice('ハーネスでは SharePoint の編集画面を開けません。実際にはライブラリのビューを新しいタブで開きます。');
    return;
  }

  if (action === 'open-app') {
    openInApp();
  }
}

function findAction(start: Node | null): ToolbarAction | undefined {
  let node: Node | null = start;

  while (node && node !== root) {
    if (node.nodeType === 1) {
      const value: string = (node as Element).getAttribute(ACTION_ATTRIBUTE) || '';
      if (value !== '') {
        return value as ToolbarAction;
      }
    }
    node = node.parentNode;
  }

  return undefined;
}

// -------------------------------------------------------------------- wiring

root.addEventListener('dblclick', (event: MouseEvent): void => {
  const action: ToolbarAction | undefined = findAction(event.target as Node | null);

  if (action === 'outline') {
    event.preventDefault();
    if (outlineClickTimer) {
      window.clearTimeout(outlineClickTimer);
      outlineClickTimer = 0;
    }
    pinPanel(pinnedPanel === 'outline' ? undefined : 'outline');
    return;
  }

  if (action === 'menu') {
    event.preventDefault();
    if (menuClickTimer) {
      window.clearTimeout(menuClickTimer);
      menuClickTimer = 0;
    }
    if (pinnedPanel === 'tree') {
      pinPanel(undefined);
    } else {
      navPanelOpen = false;
      navHost.innerHTML = '';
      pinPanel('tree');
    }
  }
});

root.addEventListener('click', (event: MouseEvent): void => {
  if (!isPlainLeftClick(event)) {
    return;
  }

  const action: ToolbarAction | undefined = findAction(event.target as Node | null);
  if (action) {
    event.preventDefault();
    runAction(action);
    return;
  }

  const link: string = findDocLinkPath(event.target as Node | null, root);
  // Folders are destinations too now, so a link to one has to survive this.
  const path: string = sanitizeDocPath(link, true);
  if (path !== '') {
    event.preventDefault();
    void navigate(path, true, splitLink(link).hash);
    return;
  }

  const code: string = findCodeToCopy(event.target as Node | null, root);
  if (code !== '') {
    event.preventDefault();
    void navigator.clipboard.writeText(code).catch((): void => undefined);
    showToast({ container: toasts, message: 'コードをコピーしました' });
    return;
  }

  const anchor: string = findHeadingAnchor(event.target as Node | null, root);
  if (anchor !== '') {
    event.preventDefault();
    const url: string = window.location.origin + buildDocUrl(currentPath) + '#' + anchor;
    void navigator.clipboard.writeText(url).catch((): void => undefined);
    showToast({ container: toasts, message: 'リンクをコピーしました' });
    return;
  }

  if (toggleHeadingAt(event.target as Node | null, content)) {
    event.preventDefault();
  }
});

document.getElementById('nav')?.addEventListener('click', (event: MouseEvent): void => {
  const target: HTMLElement = event.target as HTMLElement;
  const path: string | null = target.getAttribute('data-path');
  if (!path || !isPlainLeftClick(event)) {
    return;
  }
  event.preventDefault();
  void navigate(path, true);
});

window.addEventListener('popstate', (): void => {
  const tag: string = getQueryParam(window.location.search, TAG_PARAMETER);
  if (tag !== '') {
    searchTag(tag, false);
    return;
  }

  // Always re-render, even for the path already in `currentPath`: coming back
  // from a tag search lands on the same document with different content on
  // screen.
  void navigate(readPathFromUrl(), false);
});

void loadMetadata()
  .catch((): void => {
    /* the viewer still works without tags */
  })
  .then((): void => {
    const tag: string = getQueryParam(window.location.search, TAG_PARAMETER);
    if (tag !== '') {
      drawToolbar(false);
      searchTag(tag, false);
      return;
    }
    void navigate(readPathFromUrl(), false);
  });
