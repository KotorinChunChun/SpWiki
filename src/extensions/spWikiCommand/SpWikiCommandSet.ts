import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'SpWikiCommandSetStrings';

import { encodePath, isDocumentPath, joinUrl } from '../../webparts/spWiki/core/pathUtils';
import { DEFAULT_APP_SCHEME, buildProtocolUrl } from '../../webparts/spWiki/core/appLink';
import { setQueryParam } from '../../webparts/spWiki/core/urlUtils';
import { SharePointDocumentProvider } from '../../webparts/spWiki/services/SharePointDocumentProvider';
import MarkdownPreviewDialog from './MarkdownPreviewDialog';
import NoticeDialog from './NoticeDialog';

export interface ISpWikiCommandSetProperties {
  /**
   * Page hosting the SpWiki web part. Absolute, server relative
   * (`/sites/docs/SitePages/SpWiki.aspx`) or just a file name, which is resolved
   * against the current site's SitePages folder.
   */
  viewerPageUrl?: string;

  /**
   * Server relative library root the `?file=` paths are relative to, e.g.
   * `/sites/docs/Shared Documents`. Empty means the library the command was
   * invoked from, which is what you want unless the web part is pointed at a
   * sub-folder of that library.
   */
  libraryRootUrl?: string;

  /**
   * Value of the undocumented `env` parameter appended to the viewer URL:
   * `Embedded` hides the site chrome, `WebView` also hides hub navigation.
   * Empty adds nothing.
   */
  chromeMode?: string;

  /** Open the full page viewer in a new tab instead of the current one. */
  openInNewTab?: boolean;

  /**
   * URL scheme registered by dev/scripts/open-in-app, used by the "edit in desktop
   * app" command. Defaults to `spmd`.
   */
  appProtocol?: string;

  /** Set to false to hide the corresponding command. */
  showFullPageCommand?: boolean;
  showPreviewCommand?: boolean;
  showOpenInAppCommand?: boolean;
}

const COMMAND_FULLPAGE: string = 'SPWIKI_FULLPAGE';
const COMMAND_PREVIEW: string = 'SPWIKI_PREVIEW';
const COMMAND_OPENAPP: string = 'SPWIKI_OPENAPP';
const QUERY_PARAMETER: string = 'file';

/** How long to wait for the desktop application before offering the fallback. */
const PROTOCOL_TIMEOUT_MS: number = 1500;

/**
 * Adds "open in SpWiki" commands to a document library toolbar,
 * so a Markdown or HTML document can be read straight from the library —
 * either full page, or in a modal over the library with no page involved.
 */
export default class SpWikiCommandSet extends BaseListViewCommandSet<ISpWikiCommandSetProperties> {
  public onInit(): Promise<void> {
    this._applyCommandTitles();
    this._updateCommandVisibility();
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  public onDispose(): void {
    this.context.listView.listViewStateChangedEvent.remove(this, this._onListViewStateChanged);
    super.onDispose();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const path: string = this._getSelectedPath();
    if (path === '') {
      return;
    }

    if (event.itemId === COMMAND_FULLPAGE) {
      this._openFullPage(path);
      return;
    }

    if (event.itemId === COMMAND_PREVIEW) {
      this._openPreview(path);
      return;
    }

    if (event.itemId === COMMAND_OPENAPP) {
      this._openInApp(path);
    }
  }

  // ------------------------------------------------------------------ commands

  private _onListViewStateChanged = (): void => {
    this._updateCommandVisibility();
    this.raiseOnChange();
  };

  /**
   * SharePoint showed the manifest's `default` title even on a Japanese site,
   * ignoring the `ja-jp` entry next to it, so the labels are taken from the
   * localized resource instead — which does resolve correctly.
   */
  private _applyCommandTitles(): void {
    const fullPage: Command | undefined = this.tryGetCommand(COMMAND_FULLPAGE);
    if (fullPage) {
      fullPage.title = strings.FullPageCommandTitle;
    }

    const preview: Command | undefined = this.tryGetCommand(COMMAND_PREVIEW);
    if (preview) {
      preview.title = strings.PreviewCommandTitle;
    }

    const openApp: Command | undefined = this.tryGetCommand(COMMAND_OPENAPP);
    if (openApp) {
      openApp.title = strings.OpenInAppCommandTitle;
    }
  }

  private _updateCommandVisibility(): void {
    const enabled: boolean = this._getSelectedPath() !== '';

    const fullPage: Command | undefined = this.tryGetCommand(COMMAND_FULLPAGE);
    if (fullPage) {
      fullPage.visible = enabled && this.properties.showFullPageCommand !== false;
    }

    const preview: Command | undefined = this.tryGetCommand(COMMAND_PREVIEW);
    if (preview) {
      preview.visible = enabled && this.properties.showPreviewCommand !== false;
    }

    const openApp: Command | undefined = this.tryGetCommand(COMMAND_OPENAPP);
    if (openApp) {
      openApp.visible = enabled && this.properties.showOpenInAppCommand !== false;
    }
  }

  private _openFullPage(path: string): void {
    const url: string = this._buildViewerUrl(path);
    if (url === '') {
      // Saying so rather than returning quietly. A command that answers a click
      // with nothing at all reads as a broken deployment, and cost a real
      // afternoon: the extension was loaded and the command was on screen, but
      // `viewerPageUrl` had not reached it yet, so pressing it did nothing and
      // there was no way to tell that from the outside.
      this._showNotice(strings.NotConfiguredMessage);
      return;
    }

    if (this.properties.openInNewTab) {
      window.open(url, '_blank', 'noopener,noreferrer');
    } else {
      window.location.href = url;
    }
  }

  private _openPreview(path: string): void {
    const libraryRoot: string = this._getLibraryRootUrl();
    const provider: SharePointDocumentProvider = new SharePointDocumentProvider(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl,
      this._toLibraryName(libraryRoot)
    );

    const dialog: MarkdownPreviewDialog = new MarkdownPreviewDialog({
      path: path,
      provider: provider,
      libraryServerRelativeUrl: libraryRoot,
      buildFullPageUrl: (docPath: string): string => this._buildViewerUrl(docPath)
    });

    dialog.show().catch((): void => undefined);
  }

  /**
   * Hands the file to whatever application is registered for the scheme,
   * exactly like the web part's "open in app" button: the file's own URL with
   * the scheme swapped, and nothing else.
   */
  private _openInApp(path: string): void {
    const libraryRoot: string = this._getLibraryRootUrl();
    const fileUrl: string = this._getOrigin() + encodePath(joinUrl(libraryRoot, path));
    const scheme: string = (this.properties.appProtocol || DEFAULT_APP_SCHEME).replace(/:.*$/, '');

    // An unregistered scheme fails silently in every browser, so the only way
    // to tell is to check whether we still have focus a moment later.
    window.setTimeout((): void => {
      if (document.hasFocus()) {
        this._showNotice(strings.OpenInAppFailedMessage.replace('{0}', scheme));
      }
    }, PROTOCOL_TIMEOUT_MS);

    try {
      window.location.href = buildProtocolUrl(scheme, fileUrl);
    } catch {
      this._showNotice(strings.OpenInAppFailedMessage.replace('{0}', scheme));
    }
  }

  /** A message with nothing to copy — a misconfiguration the admin has to fix. */
  private _showNotice(message: string): void {
    const dialog: NoticeDialog = new NoticeDialog({
      message: message,
      closeLabel: strings.CloseLabel
    });

    dialog.show().catch((): void => undefined);
  }

  // ------------------------------------------------------------------- paths

  /**
   * Library relative path of the single selected file, or '' when the selection
   * is not one file the viewer can render (`.md` / `.markdown` / `.html`).
   */
  private _getSelectedPath(): string {
    const rows = this.context.listView.selectedRows;
    if (!rows || rows.length !== 1) {
      return '';
    }

    const fileRef: string = String(rows[0].getValueByName('FileRef') || '');
    if (fileRef === '' || !isDocumentPath(fileRef)) {
      return '';
    }

    const root: string = this._getLibraryRootUrl();
    const prefix: string = root + '/';

    return fileRef.slice(0, prefix.length).toLowerCase() === prefix.toLowerCase()
      ? fileRef.slice(prefix.length)
      : '';
  }

  /** Server relative URL the `?file=` paths are relative to. */
  private _getLibraryRootUrl(): string {
    if (this.properties.libraryRootUrl) {
      return this.properties.libraryRootUrl.replace(/\/+$/, '');
    }

    const list = this.context.pageContext.list;
    return list ? list.serverRelativeUrl.replace(/\/+$/, '') : '';
  }

  private _getOrigin(): string {
    const match: RegExpExecArray | null = /^(https?:\/\/[^/]+)/i.exec(
      this.context.pageContext.web.absoluteUrl
    );
    return match ? match[1] : '';
  }

  /** The same folder expressed relative to the site, for the document provider. */
  private _toLibraryName(libraryRootUrl: string): string {
    const web: string = this.context.pageContext.web.serverRelativeUrl.replace(/\/+$/, '');

    return libraryRootUrl.slice(0, web.length).toLowerCase() === web.toLowerCase()
      ? libraryRootUrl.slice(web.length).replace(/^\/+/, '')
      : libraryRootUrl.replace(/^\/+/, '');
  }

  /** Full page viewer URL for a document, or '' when no page is configured. */
  private _buildViewerUrl(path: string): string {
    const page: string = this._getViewerPagePath();
    if (page === '') {
      return '';
    }

    let search: string = setQueryParam('', QUERY_PARAMETER, path);
    if (this.properties.chromeMode) {
      search = setQueryParam(search, 'env', this.properties.chromeMode);
    }

    return page + search;
  }

  private _getViewerPagePath(): string {
    const configured: string = (this.properties.viewerPageUrl || '').replace(/\?.*$/, '');
    if (configured === '') {
      return '';
    }

    if (/^https?:\/\//i.test(configured) || configured.charAt(0) === '/') {
      return configured;
    }

    const web: string = this.context.pageContext.web.serverRelativeUrl.replace(/\/+$/, '');
    return web + '/SitePages/' + configured;
  }
}
