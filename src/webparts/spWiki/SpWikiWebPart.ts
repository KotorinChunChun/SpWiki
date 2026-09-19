import { Version } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

// Shared with the list view command set's preview dialog, so both surfaces
// render a document identically.
import styles from '../../styles/viewer.module.scss';
import * as strings from 'SpWikiWebPartStrings';

import { renderDocument } from './core/renderer';
import { sanitizeHtml } from './core/sanitize';
import {
  dirname,
  isDocumentPath,
  isFolderPath,
  normalizePath,
  sanitizeDocPath,
  splitLink
} from './core/pathUtils';
import { getQueryParam, removeQueryParam, setQueryParam } from './core/urlUtils';
import { findDocLinkPath, isPlainLeftClick } from './core/domRewrite';
import { findHeadingAnchor, toggleHeadingAt } from './core/headings';
import { findCodeToCopy } from './core/codeBlocks';
import { buildHeadingNav } from './core/toc';
import { IDirective, findDirectives, renderDirectives } from './core/directives';
import { BUNDLED_DOCS } from './core/bundledDocs';
import { findHelpLink, renderHelpPage } from './core/helpPage';
import {
  IPermissionHandle,
  IPermissionState,
  IPrincipal,
  IRoleDefinition,
  renderPermissionPanel
} from './core/permissions';
import { SharePointPermissionService } from './services/SharePointPermissionService';
import { ITreeFolder, buildTree, isEmptyTree, renderTree } from './core/pageTree';
import {
  ISearchHandle,
  ISearchHit,
  ISearchRequest,
  extractSnippet,
  matchByName,
  onlyRestricted,
  pagesInScope,
  renderSearchPage
} from './core/search';
import { DOC_LINK_ATTRIBUTE, IDocLinkContext, IDocumentResult } from './core/types';
import { ITagFieldInfo, ITagGroup, flattenTagGroups, mergeTagGroups } from './core/tags';
import {
  ITagFieldCacheState,
  canPickTagField,
  shouldLoadTagFields,
  tagFieldStatus
} from './core/tagFieldState';
import { IRelatedPage, filterPagesByTag, renderTagSearch } from './core/relatedPages';
import {
  INewPageRequest,
  buildAncestors,
  renderNavPanel,
  renderRenamePanel,
  withExtension
} from './core/navPanel';
import { newDocumentTemplate } from './core/templates';
import {
  ACTION_ATTRIBUTE,
  ITagEditorHandle,
  IToolbarButton,
  ToolbarAction,
  renderNotice,
  renderTagBar,
  renderTagEditor,
  showToast,
  renderToolbar
} from './core/chrome';
import { DEFAULT_APP_SCHEME, buildProtocolUrl } from './core/appLink';
import { LIBRARY_NOT_FOUND, SharePointDocumentProvider } from './services/SharePointDocumentProvider';
import {
  IDocumentMetadata,
  IFolderListing,
  SharePointLibraryService
} from './services/SharePointLibraryService';

export interface ISpWikiWebPartProps {
  /**
   * Name of this viewer, shown on the top row of the header. Clicking it goes
   * back to `startFile`, so it doubles as the way home.
   */
  title: string;
  /** Library title (`ドキュメント`) or folder path (`Shared Documents`). */
  libraryName: string;
  /** Document shown when the page URL carries no `?file=` parameter. */
  startFile: string;
  /** Absolute site URL; empty means "the current site". */
  siteUrl: string;

  /** Internal name of the library column holding the tags. */
  tagFieldName: string;
  /** Second tag column; empty means unused. */
  tagFieldName2: string;
  /** Third tag column; empty means unused. */
  tagFieldName3: string;
  /**
   * Page that lists the pages carrying a tag. Empty searches on this page,
   * which needs no setup at all; a value (e.g. `search_tag.aspx`) sends the
   * reader to a page of its own.
   */
  tagSearchPageUrl: string;
  /** Show the button that opens the file in SharePoint. */
  showEditButton: boolean;
  /** Show the button that hands the file to a desktop application. */
  showOpenInAppButton: boolean;
  /** Show the button that opens the tag editor. */
  showTagEditor: boolean;
  /** Show the button that opens the permissions panel. */
  showPermissionsButton: boolean;
  /** URL scheme an application on the reader's machine is registered for. */
  appProtocol: string;
}

/** Name of the query string parameter that carries the document path. */
const QUERY_PARAMETER: string = 'file';

/** Query string parameter that turns the web part into the tag search. */
const TAG_PARAMETER: string = 'tag';

/** Query string parameters that turn the web part into the search page. */
const SEARCH_PARAMETER: string = 'q';
const SEARCH_MODE_PARAMETER: string = 'qmode';
const SEARCH_SCOPE_PARAMETER: string = 'qin';
const SEARCH_RESTRICTED_PARAMETER: string = 'qlocked';

/** Query string parameter that shows the help page. */
const HELP_PARAMETER: string = 'help';

/** Undocumented parameter SharePoint reads to decide how much chrome to draw. */
const CHROME_PARAMETER: string = 'env';

/**
 * How many documents a body search reads at once.
 *
 * A body search is one request per file. Firing them all at once is what makes
 * SharePoint start refusing them, and doing them one at a time makes a library
 * of any size unusable; six keeps both ends happy.
 */
const SEARCH_CONCURRENCY: number = 6;

/**
 * How long a single click waits to see whether a second one follows.
 * Long enough for a deliberate double click, short enough not to feel laggy.
 */
const DOUBLE_CLICK_DELAY_MS: number = 250;

/**
 * What the left pane is showing.
 *
 * One pane, one occupant: pinning the second replaces the first. Stacking them
 * would halve the space each gets and leave the document narrower still, and
 * the two answer the same question — "where am I" — from different angles.
 */
type SidePanel = 'outline' | 'tree';

/** The `env` value that hides the site navigation as well as the suite bar. */
const FULL_SCREEN_CHROME: string = 'WebView';

/** How long to wait for the desktop application before offering the fallback. */
const PROTOCOL_TIMEOUT_MS: number = 1500;

/** Settle time before reading the library's columns, so typing does not spam it. */
const TAG_FIELD_LOAD_DELAY_MS: number = 400;

export default class SpWikiWebPart extends BaseClientSideWebPart<ISpWikiWebPartProps> {
  private _provider: SharePointDocumentProvider | undefined;
  private _library: SharePointLibraryService | undefined;

  private _rootElement: HTMLElement | undefined;
  private _toolbarElement: HTMLElement | undefined;
  private _tagBarElement: HTMLElement | undefined;
  private _navElement: HTMLElement | undefined;
  private _noticeElement: HTMLElement | undefined;
  private _tagPanelElement: HTMLElement | undefined;
  private _renamePanelElement: HTMLElement | undefined;

  private _permPanelElement: HTMLElement | undefined;
  private _toastElement: HTMLElement | undefined;
  private _contentElement: HTMLElement | undefined;
  private _bodyElement: HTMLElement | undefined;
  private _outlinePopupElement: HTMLElement | undefined;
  private _sidePaneElement: HTMLElement | undefined;
  private _topPanelElement: HTMLElement | undefined;
  private _footerPathElement: HTMLElement | undefined;

  private _currentPath: string = '';
  private _metadata: IDocumentMetadata | undefined;
  private _tagFields: ITagFieldInfo[] = [];
  private _tagFieldOptions: IPropertyPaneDropdownOption[] = [];
  private _tagFieldsLoaded: boolean = false;
  private _tagFieldsLoading: boolean = false;
  /** Library the cached columns were read from; a change makes them stale. */
  private _tagFieldsKey: string = '';
  private _tagFieldsTimer: number = 0;
  /** Folders of the library, for the rename panel's suggestions. */
  private _folders: string[] = [];
  private _tagPanelOpen: boolean = false;
  /** Internal name of the column the open editor is for. */
  private _tagPanelField: string = '';
  /** Handle on the open editor, used to report a failed save inside it. */
  private _tagEditor: ITagEditorHandle | undefined;
  private _navPanelOpen: boolean = false;
  private _renamePanelOpen: boolean = false;
  private _helpOpen: boolean = false;
  /** The search page is showing instead of a document. */
  private _searchOpen: boolean = false;
  private _searchHandle: ISearchHandle | undefined;
  /** Occupant of the left pane; undefined when the pane is closed. */
  private _pinnedPanel: SidePanel | undefined;
  /** The outline popup, which is separate from the pinned pane. */
  private _outlinePopupOpen: boolean = false;
  private _outlineClickTimer: number = 0;
  private _menuClickTimer: number = 0;
  private _permissions: SharePointPermissionService | undefined;
  private _permPanelOpen: boolean = false;
  private _permHandle: IPermissionHandle | undefined;
  /**
   * Drives the padlock in the header. Undefined until the item metadata has
   * been read, which is what the icon's neutral state means.
   */
  private _hasUniquePermissions: boolean | undefined;

  /** Guards against an earlier, slower request overwriting a newer one. */
  private _loadToken: number = 0;

  public render(): void {
    this._renderShell();

    if (!this.properties.libraryName) {
      this._showStatus(styles.status, strings.NotConfiguredMessage);
      return;
    }

    this._provider = new SharePointDocumentProvider(
      this.context.spHttpClient,
      this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
      this.properties.libraryName
    );
    this._library = undefined;

    // Only throw the cached columns away when they came from a different
    // library. `render()` runs on every property change, and clearing them
    // unconditionally used to leave the tag dropdowns disabled for good: the
    // reload only ever ran when the property pane was opened.
    if (this._tagFieldsKey !== this._getLibraryKey()) {
      this._tagFields = [];
      this._tagFieldOptions = [];
      this._tagFieldsLoaded = false;
    }

    if (getQueryParam(window.location.search, HELP_PARAMETER) !== '' || this._helpOpen) {
      this._renderToolbar();
      this._openHelp(getQueryParam(window.location.search, HELP_PARAMETER), false);
      return;
    }

    // `?q=` turns the same web part into the search page, the same way `?tag=`
    // turns it into the tag search: a page dedicated to it needs the web part
    // and the parameter, and nothing else.
    if (getQueryParam(window.location.search, SEARCH_PARAMETER) !== '' || this._searchOpen) {
      this._renderToolbar();
      this._openSearch({
        term: getQueryParam(window.location.search, SEARCH_PARAMETER),
        mode:
          getQueryParam(window.location.search, SEARCH_MODE_PARAMETER) === 'content'
            ? 'content'
            : 'name',
        scope: getQueryParam(window.location.search, SEARCH_SCOPE_PARAMETER),
        restrictedOnly:
          getQueryParam(window.location.search, SEARCH_RESTRICTED_PARAMETER) === '1'
      });
      return;
    }

    // `?tag=` turns the same web part into the tag search, so a page dedicated
    // to it (search_tag.aspx) needs nothing but the parameter.
    const tag: string = getQueryParam(window.location.search, TAG_PARAMETER);
    if (tag !== '') {
      this._renderToolbar();
      void this._searchTag(tag, false);
      return;
    }

    void this._navigate(this._readPathFromUrl(), false);
  }

  protected onInit(): Promise<void> {
    window.addEventListener('popstate', this._onPopState);
    return Promise.resolve();
  }

  protected onDispose(): void {
    if (this._tagFieldsTimer) {
      window.clearTimeout(this._tagFieldsTimer);
      this._tagFieldsTimer = 0;
    }
    window.removeEventListener('popstate', this._onPopState);
    if (this._rootElement) {
      this._rootElement.removeEventListener('click', this._onRootClick);
      this._rootElement.removeEventListener('dblclick', this._onRootDoubleClick);
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.1');
  }

  // ----------------------------------------------------------- property pane

  /**
   * The tag column dropdowns are filled from the library itself, so the list
   * has to be fetched before the pane can render its options.
   */
  protected onPropertyPaneConfigurationStart(): void {
    this._ensureTagFieldsLoaded();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    // A different library or column invalidates everything cached from it. The
    // reload itself is left to `_ensureTagFieldsLoaded`, which the pane calls
    // on every render.
    if (propertyPath === 'libraryName' || propertyPath === 'siteUrl') {
      this._tagFields = [];
      this._tagFieldOptions = [];
      this._tagFieldsLoaded = false;
    }
    if (propertyPath.indexOf('tagFieldName') === 0 && this._library) {
      this._library.invalidatePages();
    }
  }

  /** The cache state the rules in `core/tagFieldState` are written against. */
  private _getTagFieldState(): ITagFieldCacheState {
    return {
      libraryKey: this._getLibraryKey(),
      loadedKey: this._tagFieldsKey,
      loaded: this._tagFieldsLoaded,
      loading: this._tagFieldsLoading,
      hasLibrary: !!this.properties.libraryName,
      optionCount: this._tagFieldOptions.length
    };
  }

  /** Why the tag column dropdowns look the way they do. */
  private _getTagFieldStatus(): string {
    switch (tagFieldStatus(this._getTagFieldState())) {
      case 'needs-library':
        return strings.TagFieldNeedsLibrary;
      case 'loading':
        return strings.TagFieldLoading;
      case 'none-found':
        return strings.TagFieldNoneFound;
      default:
        return strings.TagFieldReady.replace('{0}', String(this._tagFieldOptions.length));
    }
  }

  /** Identifies the library the cached columns belong to. */
  private _getLibraryKey(): string {
    return (this.properties.siteUrl || '') + '|' + (this.properties.libraryName || '');
  }

  /**
   * Reads the library's columns unless that has already been done for the
   * library currently configured.
   *
   * Called from every property pane render rather than only from
   * `onPropertyPaneConfigurationStart`, because the pane stays open while the
   * library name is being typed: without this the dropdowns would be disabled
   * until the pane was closed and opened again.
   *
   * The delay is what keeps a reactive pane from firing one request per
   * keystroke while the library name is typed.
   */
  private _ensureTagFieldsLoaded(): void {
    if (!shouldLoadTagFields(this._getTagFieldState())) {
      return;
    }

    const key: string = this._getLibraryKey();

    if (this._tagFieldsTimer) {
      window.clearTimeout(this._tagFieldsTimer);
    }

    this._tagFieldsTimer = window.setTimeout((): void => {
      this._tagFieldsTimer = 0;
      this._loadTagFieldsNow(key);
    }, TAG_FIELD_LOAD_DELAY_MS);
  }

  private _loadTagFieldsNow(key: string): void {
    this._tagFieldsLoading = true;

    this._loadTagFields()
      .catch((): void => {
        // An unreadable library leaves the list empty; the pane explains it
        // rather than staying greyed out with no reason given.
        this._tagFields = [];
        this._tagFieldOptions = [];
      })
      .then((): void => {
        this._tagFieldsLoading = false;
        this._tagFieldsLoaded = true;
        this._tagFieldsKey = key;
        this.context.propertyPane.refresh();
      })
      .catch((): void => undefined);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // The pane calls this on every render, which is what keeps the column list
    // arriving even when the pane was opened before the library was named.
    this._ensureTagFieldsLoaded();

    const noneOption: IPropertyPaneDropdownOption = { key: '', text: strings.TagFieldNone };
    const tagOptions: IPropertyPaneDropdownOption[] = [noneOption].concat(this._tagFieldOptions);
    const tagFieldsReady: boolean = canPickTagField(this._getTagFieldState());

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldDescription
                }),
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlFieldLabel,
                  description: strings.SiteUrlFieldDescription
                }),
                PropertyPaneTextField('libraryName', {
                  label: strings.LibraryNameFieldLabel,
                  description: strings.LibraryNameFieldDescription
                }),
                PropertyPaneTextField('startFile', {
                  label: strings.StartFileFieldLabel,
                  description: strings.StartFileFieldDescription
                })
              ]
            },
            {
              groupName: strings.TagGroupName,
              groupFields: [
                PropertyPaneLabel('tagFieldStatus', { text: this._getTagFieldStatus() }),
                PropertyPaneDropdown('tagFieldName', {
                  label: strings.TagFieldLabel,
                  options: tagOptions,
                  disabled: !tagFieldsReady
                }),
                PropertyPaneDropdown('tagFieldName2', {
                  label: strings.TagFieldLabel2,
                  options: tagOptions,
                  disabled: !tagFieldsReady
                }),
                PropertyPaneDropdown('tagFieldName3', {
                  label: strings.TagFieldLabel3,
                  options: tagOptions,
                  disabled: !tagFieldsReady
                }),
                PropertyPaneTextField('tagSearchPageUrl', {
                  label: strings.TagSearchPageLabel,
                  description: strings.TagSearchPageDescription
                }),
                PropertyPaneToggle('showPermissionsButton', {
                  label: strings.ShowPermissionsButtonLabel,
                  checked: this._isEnabled('showPermissionsButton')
                }),
                PropertyPaneToggle('showTagEditor', {
                  label: strings.ShowTagEditorLabel,
                  checked: this._isEnabled('showTagEditor')
                })
              ]
            },
            {
              groupName: strings.ToolbarGroupName,
              groupFields: [
                PropertyPaneToggle('showEditButton', {
                  label: strings.ShowEditButtonLabel,
                  checked: this._isEnabled('showEditButton')
                }),
                PropertyPaneToggle('showOpenInAppButton', {
                  label: strings.ShowOpenInAppButtonLabel,
                  checked: this._isEnabled('showOpenInAppButton')
                }),
                PropertyPaneTextField('appProtocol', {
                  label: strings.AppProtocolLabel,
                  description: strings.AppProtocolDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /** アプリ連携は明示的な有効化が必要。他の表示設定は既定で有効。 */
  private _isEnabled(name: keyof ISpWikiWebPartProps): boolean {
    if (name === 'showOpenInAppButton') {
      return this.properties.showOpenInAppButton === true;
    }
    return this.properties[name] !== false;
  }

  /** Fills the column cache. The flags around it belong to the caller. */
  private async _loadTagFields(): Promise<void> {
    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library) {
      this._tagFields = [];
      this._tagFieldOptions = [];
      return;
    }

    this._tagFields = await library.getTagFields();
    this._tagFieldOptions = this._tagFields.map((field: ITagFieldInfo): IPropertyPaneDropdownOption => {
      return { key: field.internalName, text: field.title + ' (' + field.internalName + ')' };
    });
  }

  // ---------------------------------------------------------------- rendering

  private _renderShell(): void {
    if (this._rootElement) {
      this._rootElement.removeEventListener('click', this._onRootClick);
      this._rootElement.removeEventListener('dblclick', this._onRootDoubleClick);
    }

    // The rename panel sits directly under the file name it is about, above the
    // tags — editing the name from a box three sections below it read as though
    // it belonged to the tags.
    this.domElement.innerHTML = `
      <div class="${styles.spWiki}">
        <div class="${styles.topPanel}">
          <div class="${styles.header}"></div>
          <div class="${styles.renamePanel}"></div>
          <div class="${styles.tagBar}"></div>
          <div class="${styles.nav}"></div>
          <div class="${styles.outlinePopup}"></div>
        </div>
        <div class="${styles.permPanel}"></div>
        <div class="${styles.notice}"></div>
        <div class="${styles.tagPanel}"></div>
        <div class="${styles.body}">
          <div class="${styles.sidePane}"></div>
          <div class="${styles.content}"></div>
        </div>
        <div class="${styles.footer}">${escape(strings.PageLabel)} <span></span></div>
        <div class="${styles.toasts}"></div>
      </div>`;

    this._rootElement = this.domElement.querySelector('.' + styles.spWiki) as HTMLElement;
    this._toolbarElement = this.domElement.querySelector('.' + styles.header) as HTMLElement;
    this._renamePanelElement = this.domElement.querySelector('.' + styles.renamePanel) as HTMLElement;

    this._permPanelElement = this.domElement.querySelector('.' + styles.permPanel) as HTMLElement;
    this._tagBarElement = this.domElement.querySelector('.' + styles.tagBar) as HTMLElement;
    this._navElement = this.domElement.querySelector('.' + styles.nav) as HTMLElement;
    this._noticeElement = this.domElement.querySelector('.' + styles.notice) as HTMLElement;
    this._tagPanelElement = this.domElement.querySelector('.' + styles.tagPanel) as HTMLElement;
    this._contentElement = this.domElement.querySelector('.' + styles.content) as HTMLElement;
    this._bodyElement = this.domElement.querySelector('.' + styles.body) as HTMLElement;
    this._outlinePopupElement = this.domElement.querySelector('.' + styles.outlinePopup) as HTMLElement;
    this._sidePaneElement = this.domElement.querySelector('.' + styles.sidePane) as HTMLElement;
    this._topPanelElement = this.domElement.querySelector('.' + styles.topPanel) as HTMLElement;
    this._footerPathElement = this.domElement.querySelector('.' + styles.footer + ' span') as HTMLElement;
    this._toastElement = this.domElement.querySelector('.' + styles.toasts) as HTMLElement;

    this._rootElement.addEventListener('click', this._onRootClick);
    this._rootElement.addEventListener('dblclick', this._onRootDoubleClick);
  }

  private _renderToolbar(): void {
    if (!this._toolbarElement) {
      return;
    }

    const buttons: IToolbarButton[] = [];
    const ready: boolean = this._metadata !== undefined;

    if (this._isEnabled('showEditButton')) {
      buttons.push({ action: 'edit', label: strings.EditButtonLabel, disabled: !ready });
    }
    if (this._isEnabled('showOpenInAppButton')) {
      buttons.push({ action: 'open-app', label: strings.OpenInAppButtonLabel, disabled: !ready });
    }

    renderToolbar({
      container: this._toolbarElement,
      path: this._currentPath,
      buttons: buttons,
      menuLabel: strings.MenuButtonLabel,
      title: this.properties.title || strings.DefaultTitle,
      titleLabel: strings.HomeButtonLabel,
      rootLabel: strings.RootLabel,
      // Renaming needs the list item, so the file name stays plain text until
      // the metadata is in.
      pathLabel: ready ? strings.PathButtonLabel : undefined,
      outline: {
        label: strings.OutlineButtonLabel,
        pinned: this._pinnedPanel === 'outline'
      },
      menuPinned: this._pinnedPanel === 'tree',
      folderLinkAttribute: DOC_LINK_ATTRIBUTE,
      buildFolderUrl: (folderPath: string): string => this._buildDocUrl(folderPath),
      permission: this._isEnabled('showPermissionsButton')
        ? {
            locked: this._hasUniquePermissions,
            label:
              this._hasUniquePermissions === undefined
                ? strings.PermUnknownTitle
                : this._hasUniquePermissions
                  ? strings.PermLockedTitle
                  : strings.PermUnlockedTitle
          }
        : undefined
    });
  }

  /** Internal names of the configured tag columns, in order, without blanks. */
  private _getTagFieldNames(): string[] {
    const names: string[] = [];
    const configured: string[] = [
      this.properties.tagFieldName || '',
      this.properties.tagFieldName2 || '',
      this.properties.tagFieldName3 || ''
    ];

    for (let i: number = 0; i < configured.length; i++) {
      if (configured[i] !== '' && names.indexOf(configured[i]) < 0) {
        names.push(configured[i]);
      }
    }

    return names;
  }

  private _renderTagBar(): void {
    if (!this._tagBarElement) {
      return;
    }

    if (!this._metadata) {
      this._tagBarElement.innerHTML = '';
      return;
    }

    renderTagBar({
      container: this._tagBarElement,
      groups: this._metadata.tagGroups,
      buildTagUrl: (tag: string): string => this._buildTagUrl(tag),
      onTag: (tag: string): void => {
        void this._searchTag(tag, true);
      },
      // The column name is where the editor lives now, so the toolbar has no
      // tag button of its own. Clicking the column that is already open closes
      // it; clicking a different one switches to that column.
      onEdit: this._isEnabled('showTagEditor')
        ? (internalName: string): void => {
            if (this._tagPanelOpen && this._tagPanelField === internalName) {
              this._closeTagPanel();
            } else {
              void this._openTagPanel(internalName);
            }
          }
        : undefined,
      editLabel: strings.TagEditLabel
    });
  }

  private async _navigate(path: string, pushHistory: boolean, hash?: string): Promise<void> {
    if (!this._provider || !this._contentElement) {
      return;
    }

    this._searchOpen = false;
    this._searchHandle = undefined;
    this._closePanels();
    this._clearNotice();
    this._clearTagBar();

    if (path === '') {
      this._setPathLabels('');
      this._showStatus(styles.status, strings.NotConfiguredMessage);
      return;
    }

    this._currentPath = path;
    this._metadata = undefined;
    this._hasUniquePermissions = undefined;
    this._setPathLabels(path);
    this._renderToolbar();
    if (pushHistory) {
      window.history.pushState(
        { md: path },
        '',
        window.location.pathname + this._buildDocSearch(path) + (hash || '')
      );
    }

    this._loadToken += 1;
    const token: number = this._loadToken;

    this._showStatus(styles.status, strings.LoadingMessage);

    // A folder has no document to fetch; its page *is* the listing.
    if (isFolderPath(path)) {
      await this._renderFolderPage(path, token);
      return;
    }

    const result: IDocumentResult = await this._provider.getDocument(path);
    if (token !== this._loadToken) {
      return; // A newer navigation won the race.
    }

    if (result.kind === 'notfound') {
      this._showNotFound(path, token);
      return;
    }

    if (result.kind === 'error') {
      const detail: string = result.message || '';
      const isLibraryError: boolean = detail.indexOf(LIBRARY_NOT_FOUND) === 0;
      this._showError(isLibraryError ? strings.LibraryNotFoundMessage : strings.ErrorMessage, detail);
      return;
    }

    // Resolved and cached by the call above, so this cannot fail here.
    const libraryUrl: string = await this._provider.getLibraryServerRelativeUrl();
    if (token !== this._loadToken) {
      return;
    }

    renderDocument({
      source: result.text || '',
      container: this._contentElement,
      sanitize: sanitizeHtml,
      context: this._createLinkContext(path, libraryUrl),
      headings: {
        copyLabel: strings.HeadingCopyLabel,
        toggleLabel: strings.HeadingToggleLabel
      },
      codeBlocks: { copyLabel: strings.CodeCopyLabel }
    });

    // The browser cannot honour a `#heading` in the address on its own: the
    // document arrives long after it gave up looking for the element.
    this._scrollToAnchor(hash !== undefined ? hash : window.location.hash);

    // The side panels are about whatever is now on screen.
    this._renderSidePanels();

    // Item metadata only drives the chrome, so the document is never held back
    // waiting for it.
    void this._loadItemState(path, token);

    // Likewise the generated lists: they need the whole library listing, and
    // the document is readable long before that arrives.
    void this._expandDirectives(path, token);
  }

  /**
   * Fills in the `## タグページ一覧:...` style headings.
   *
   * Nothing is fetched for a document that carries none, which is most of them
   * — the library listing is one request over every item, and it is not worth
   * making every page wait for it on the chance that it might be needed.
   */
  private async _expandDirectives(path: string, token: number): Promise<void> {
    if (!this._contentElement) {
      return;
    }

    const directives: IDirective[] = findDirectives(this._contentElement);
    if (directives.length === 0) {
      return;
    }

    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library || token !== this._loadToken) {
      return;
    }

    let pages: IRelatedPage[];
    try {
      pages = await library.getPages(await this._getConfiguredTagFields());
    } catch {
      // The document itself is fine; only the generated lists are missing.
      return;
    }

    if (token !== this._loadToken || !this._contentElement) {
      return;
    }

    renderDirectives(this._contentElement, directives, {
      pages: pages,
      currentPath: path,
      buildDocUrl: (docPath: string): string => this._buildDocUrl(docPath),
      linkAttribute: DOC_LINK_ATTRIBUTE,
      strings: {
        empty: strings.DirectiveEmptyMessage,
        rootLabel: strings.RootLabel
      }
    });
  }

  /** The permission service for the configured library, built once. */
  private async _ensurePermissionService(): Promise<SharePointPermissionService | undefined> {
    if (this._permissions) {
      return this._permissions;
    }

    if (!this._provider) {
      return undefined;
    }

    let libraryUrl: string;
    try {
      libraryUrl = await this._provider.getLibraryServerRelativeUrl();
    } catch {
      return undefined;
    }

    this._permissions = new SharePointPermissionService(
      this.context.spHttpClient,
      this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
      libraryUrl
    );

    return this._permissions;
  }

  /**
   * Opens the permissions panel for what is on screen.
   *
   * Reading the state costs several requests — the item, whether it is
   * inherited, who is on it, and for a folder a scan for descendants that broke
   * away — so it only happens when the reader asks for it.
   */
  private async _openPermPanel(): Promise<void> {
    if (!this._permPanelElement || this._currentPath === '') {
      return;
    }

    const service: SharePointPermissionService | undefined = await this._ensurePermissionService();
    if (!service) {
      this._showNotice(strings.LibraryNotFoundMessage);
      return;
    }

    this._permPanelOpen = true;
    this._permPanelElement.innerHTML = '';

    const loading: HTMLElement = document.createElement('p');
    loading.className = 'mdv-related-empty';
    loading.textContent = strings.PermLoadingMessage;
    this._permPanelElement.appendChild(loading);

    const path: string = this._currentPath;
    const isFolder: boolean = isFolderPath(path);

    let state: IPermissionState;
    let roles: IRoleDefinition[];
    try {
      state = await service.getState(path, isFolder);
      roles = await service.getRoleDefinitions();
    } catch (error) {
      this._permPanelElement.innerHTML = '';
      this._showNotice(strings.PermLoadFailedMessage + ' ' + describeError(error));
      return;
    }

    if (!this._permPanelOpen || !this._permPanelElement || this._currentPath !== path) {
      return; // The reader moved on while it was loading.
    }

    this._hasUniquePermissions = state.hasUnique;
    this._renderToolbar();

    this._permHandle = renderPermissionPanel({
      container: this._permPanelElement,
      state: state,
      roles: roles,
      strings: {
        title: strings.PermTitle,
        inheritedFrom: strings.PermInheritedFromLabel,
        unique: strings.PermUniqueLabel,
        team: strings.PermTeamLabel,
        folderSuffix: strings.PermFolderSuffix,
        whoHasAccess: strings.PermWhoLabel,
        restrict: strings.PermRestrictLabel,
        restore: strings.PermRestoreLabel,
        confirmTitle: strings.PermConfirmTitle,
        confirmRestrict: strings.PermConfirmRestrict,
        confirmRestore: strings.PermConfirmRestore,
        confirmYes: strings.PermConfirmYes,
        confirmNo: strings.CancelLabel,
        addLabel: strings.PermAddLabel,
        addPlaceholder: strings.PermAddPlaceholder,
        search: strings.PermSearchButton,
        searchFailed: strings.PermSearchFailedMessage,
        add: strings.PermAddButton,
        remove: strings.PermRemoveLabel,
        searching: strings.PermSearchingMessage,
        noResults: strings.PermNoResultsMessage,
        brokenTitle: strings.PermBrokenTitle,
        brokenIntro: strings.PermBrokenIntro,
        fileScopeWarning: strings.PermFileScopeWarning,
        loading: strings.PermLoadingMessage,
        close: strings.CloseLabel
      },
      searchPrincipals: (query: string): Promise<IPrincipal[]> => service.searchPrincipals(query),
      onRestrict: (): void => {
        void this._runPermChange((): Promise<void> => service.restrict(path, isFolder), true);
      },
      onRestore: (): void => {
        void this._runPermChange(
          (): Promise<void> => service.restoreInheritance(path, isFolder),
          false
        );
      },
      onAdd: (principal: IPrincipal, roleId: number): void => {
        void this._runPermChange(
          (): Promise<void> => service.grant(path, isFolder, principal, roleId)
        );
      },
      onRemove: (principal: IPrincipal): void => {
        void this._runPermChange((): Promise<void> => service.revoke(path, isFolder, principal));
      },
      onClose: (): void => this._closePermPanel()
    });
  }

  /**
   * Runs one permission change and reopens the panel on the result.
   *
   * Re-reading rather than patching what is on screen: a permission change has
   * effects the client cannot predict — breaking inheritance copies entries,
   * granting one level can replace another — and showing a guess about who can
   * see something would be worse than showing nothing.
   */
  private async _runPermChange(
    change: () => Promise<void>,
    expectUnique?: boolean
  ): Promise<void> {
    try {
      await change();
    } catch (error) {
      if (this._permHandle) {
        this._permHandle.setError(strings.PermChangeFailedMessage + ' ' + describeError(error));
      }
      return;
    }

    const path: string = this._currentPath;
    await this._openPermPanel();

    if (expectUnique === undefined || !this._permHandle || this._currentPath !== path) {
      return;
    }

    // SharePoint answered 2xx; that is not the same as the change having
    // happened. A call that succeeds and leaves the state alone used to look
    // exactly like a working one — the panel simply redrew itself unchanged,
    // with nothing to tell the reader that nothing had moved.
    const service: SharePointPermissionService | undefined = this._permissions;
    if (!service) {
      return;
    }

    let hasUnique: boolean;
    try {
      hasUnique = (await service.getState(path, isFolderPath(path))).hasUnique;
    } catch {
      return; // The panel already shows whatever could be read.
    }

    if (hasUnique !== expectUnique && this._permHandle && this._currentPath === path) {
      this._permHandle.setError(strings.PermNoEffectMessage);
    }
  }

  private _closePermPanel(): void {
    this._permPanelOpen = false;
    this._permHandle = undefined;
    if (this._permPanelElement) {
      this._permPanelElement.innerHTML = '';
    }
  }

  // ------------------------------------------------------------------ outline

  /**
   * Handles a click on the outline button.
   *
   * The single click is deferred, because a double click also fires two of
   * them: acting immediately would flash the popup open and shut on the way to
   * pinning it. The timer is the only way a browser lets the two be told apart.
   */
  private _onOutlineClick(): void {
    if (this._outlineClickTimer) {
      window.clearTimeout(this._outlineClickTimer);
    }

    this._outlineClickTimer = window.setTimeout((): void => {
      this._outlineClickTimer = 0;

      // Pinned already: the click puts the pane away rather than opening a
      // second copy of the same thing as a popup.
      if (this._pinnedPanel === 'outline') {
        this._pinPanel(undefined);
        return;
      }

      this._outlinePopupOpen = !this._outlinePopupOpen;
      this._renderSidePanels();
    }, DOUBLE_CLICK_DELAY_MS);
  }

  /** A double click pins the outline beside the document, or puts it away. */
  private _onOutlineDoubleClick(): void {
    if (this._outlineClickTimer) {
      window.clearTimeout(this._outlineClickTimer);
      this._outlineClickTimer = 0;
    }

    this._pinPanel(this._pinnedPanel === 'outline' ? undefined : 'outline');
  }

  /** Same split for the menu button, so both behave identically. */
  private _onMenuClick(folder: string): void {
    if (this._menuClickTimer) {
      window.clearTimeout(this._menuClickTimer);
    }

    this._menuClickTimer = window.setTimeout((): void => {
      this._menuClickTimer = 0;

      if (this._pinnedPanel === 'tree') {
        this._pinPanel(undefined);
        return;
      }

      if (this._navPanelOpen) {
        this._closeNavPanel();
      } else {
        void this._openNavPanel(folder);
      }
    }, DOUBLE_CLICK_DELAY_MS);
  }

  private _onMenuDoubleClick(folder: string): void {
    if (this._menuClickTimer) {
      window.clearTimeout(this._menuClickTimer);
      this._menuClickTimer = 0;
    }

    if (this._pinnedPanel === 'tree') {
      this._pinPanel(undefined);
      return;
    }

    this._closeNavPanel();
    this._pinPanel('tree');
    void this._openNavPanel(folder);
  }

  /** Puts one panel in the left pane, or empties it. */
  private _pinPanel(panel: SidePanel | undefined): void {
    this._pinnedPanel = panel;
    // A popup of the thing now pinned would be a duplicate.
    this._outlinePopupOpen = false;

    if (panel !== 'tree') {
      this._closeNavPanel();
    }

    this._renderSidePanels();
    // The buttons show which mode they are in, so the header follows.
    this._renderToolbar();
  }

  /** Draws whichever side panels the current state calls for. */
  private _renderSidePanels(): void {
    if (this._bodyElement) {
      this._bodyElement.className = this._pinnedPanel
        ? styles.body + ' ' + styles.bodyPinned
        : styles.body;
    }

    if (this._sidePaneElement && this._pinnedPanel !== 'tree') {
      this._sidePaneElement.innerHTML = '';
    }

    this._renderOutline();

    if (this._pinnedPanel === 'tree') {
      void this._openNavPanel(dirname(this._currentPath));
    }

    this._measureTopPanel();
  }

  /**
   * Draws the outline into whichever host the current mode calls for.
   *
   * Both hosts are always emptied first: switching modes has to move the
   * outline, not leave a copy behind in the one it came from.
   */
  private _renderOutline(): void {
    if (this._outlinePopupElement) {
      this._outlinePopupElement.innerHTML = '';
    }

    const pinned: boolean = this._pinnedPanel === 'outline';
    if (!pinned && !this._outlinePopupOpen) {
      return;
    }

    if (!this._contentElement) {
      return;
    }

    const host: HTMLElement | undefined = pinned
      ? this._sidePaneElement
      : this._outlinePopupElement;
    if (!host) {
      return;
    }

    if (pinned) {
      host.innerHTML = '';
    }

    const panel: HTMLElement = document.createElement('div');
    panel.className = 'mdv-outline-panel';

    const heading: HTMLElement = document.createElement('div');
    heading.className = 'mdv-outline-title';
    heading.textContent = strings.OutlineTitle;
    panel.appendChild(heading);

    const nav: HTMLElement | undefined = buildHeadingNav(this._contentElement);
    if (nav) {
      panel.appendChild(nav);
    } else {
      const empty: HTMLElement = document.createElement('p');
      empty.className = 'mdv-related-empty';
      empty.textContent = strings.OutlineEmptyMessage;
      panel.appendChild(empty);
    }

    host.appendChild(panel);
  }

  /** Renders the search page and runs the request it was opened with. */
  private _openSearch(request: ISearchRequest): void {
    if (!this._contentElement) {
      return;
    }

    this._clearView();
    this._searchOpen = true;

    this._searchHandle = renderSearchPage({
      container: this._contentElement,
      request: request,
      strings: {
        title: strings.SearchTitle,
        termPlaceholder: strings.SearchTermPlaceholder,
        modeName: strings.SearchModeName,
        modeContent: strings.SearchModeContent,
        scope: strings.SearchScopeLabel,
        scopePlaceholder: strings.SearchScopePlaceholder,
        submit: strings.SearchSubmitLabel,
        searching: strings.SearchBusyMessage,
        resultCount: strings.SearchResultCount,
        empty: strings.SearchEmptyMessage,
        prompt: strings.SearchPromptMessage,
        restrictedOnly: strings.SearchRestrictedOnlyLabel,
        locked: strings.PermLockedShortLabel
      },
      buildDocUrl: (docPath: string): string => this._buildDocUrl(docPath),
      linkAttribute: DOC_LINK_ATTRIBUTE,
      onSearch: (next: ISearchRequest): void => {
        void this._runSearch(next, true);
      }
    });

    if (request.term !== '') {
      void this._runSearch(request, false);
    }
  }

  /**
   * Runs one search.
   *
   * The address carries the request so a result set can be shared and the back
   * button works, exactly like `?file=` and `?tag=`.
   */
  private async _runSearch(request: ISearchRequest, pushHistory: boolean): Promise<void> {
    if (!this._searchHandle) {
      return;
    }

    if (pushHistory) {
      let search: string = setQueryParam('', SEARCH_PARAMETER, request.term);
      search = setQueryParam(search, SEARCH_MODE_PARAMETER, request.mode);
      if (request.scope !== '') {
        search = setQueryParam(search, SEARCH_SCOPE_PARAMETER, request.scope);
      }
      if (request.restrictedOnly) {
        search = setQueryParam(search, SEARCH_RESTRICTED_PARAMETER, '1');
      }
      window.history.pushState({ q: request.term }, '', window.location.pathname + search);
    }

    // The restricted filter answers on its own; everything else needs a term.
    if (request.term === '' && !request.restrictedOnly) {
      this._searchHandle.setResults([]);
      return;
    }

    this._loadToken += 1;
    const token: number = this._loadToken;
    this._searchHandle.setBusy(strings.SearchBusyMessage);

    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library || token !== this._loadToken) {
      return;
    }

    let pages: IRelatedPage[];
    try {
      pages = await library.getPages(await this._getConfiguredTagFields());
    } catch (error) {
      this._showNotice(strings.ErrorMessage + ' ' + describeError(error));
      return;
    }

    if (token !== this._loadToken || !this._searchHandle) {
      return;
    }

    let scoped: IRelatedPage[] = pagesInScope(pages, request.scope);
    if (request.restrictedOnly) {
      scoped = onlyRestricted(scoped);
    }

    // With no term, the filter alone is the answer.
    if (request.term === '') {
      const all: ISearchHit[] = [];
      for (let i: number = 0; i < scoped.length; i++) {
        all.push({ page: scoped[i], snippet: '' });
      }
      this._searchHandle.setResults(all);
      return;
    }

    if (request.mode === 'name') {
      this._searchHandle.setResults(matchByName(scoped, request.term));
      return;
    }

    const hits: ISearchHit[] = await this._searchContents(scoped, request.term, token);
    if (token !== this._loadToken || !this._searchHandle) {
      return;
    }

    this._searchHandle.setResults(hits);
  }

  /**
   * Reads the documents and keeps the ones containing the term.
   *
   * This is the expensive half of the search — one request per document — so it
   * runs a few at a time and reports how far it has got. A newer search
   * abandons it through the load token rather than letting two runs fight over
   * the result area.
   */
  private async _searchContents(
    pages: IRelatedPage[],
    term: string,
    token: number
  ): Promise<ISearchHit[]> {
    const provider: SharePointDocumentProvider | undefined = this._provider;
    if (!provider) {
      return [];
    }

    const hits: ISearchHit[] = [];
    let next: number = 0;
    let done: number = 0;

    const worker: () => Promise<void> = async (): Promise<void> => {
      for (;;) {
        const index: number = next;
        next += 1;

        if (index >= pages.length || token !== this._loadToken) {
          return;
        }

        const result: IDocumentResult = await provider.getDocument(pages[index].path);
        if (result.kind === 'ok') {
          const snippet: string = extractSnippet(result.text || '', term);
          if (snippet !== '') {
            hits.push({ page: pages[index], snippet: snippet });
          }
        }

        done += 1;
        if (this._searchHandle && token === this._loadToken) {
          this._searchHandle.setBusy(
            strings.SearchProgressMessage.replace('{0}', String(done)).replace(
              '{1}',
              String(pages.length)
            )
          );
        }
      }
    };

    const workers: Promise<void>[] = [];
    for (let i: number = 0; i < SEARCH_CONCURRENCY; i++) {
      workers.push(worker());
    }
    await Promise.all(workers);

    hits.sort((a: ISearchHit, b: ISearchHit): number => {
      const x: string = a.page.path.toLowerCase();
      const y: string = b.page.path.toLowerCase();
      return x < y ? -1 : x > y ? 1 : 0;
    });

    return hits;
  }

  /**
   * A folder opened as a page: the tree of everything under it.
   *
   * Folders carry list items of their own, so the tags keep working exactly as
   * they do for a document — a folder can be tagged and found by tag. What
   * changes is the body, which is the listing rather than a rendered file.
   */
  private async _renderFolderPage(path: string, token: number): Promise<void> {
    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library || token !== this._loadToken || !this._contentElement) {
      return;
    }

    let pages: IRelatedPage[];
    try {
      pages = await library.getPages(await this._getConfiguredTagFields());
    } catch (error) {
      this._showError(strings.ErrorMessage, describeError(error));
      return;
    }

    if (token !== this._loadToken || !this._contentElement) {
      return;
    }

    const tree: ITreeFolder = buildTree(pages, path);
    this._contentElement.innerHTML = '';

    const heading: HTMLElement = document.createElement('h1');
    heading.textContent = path;
    this._contentElement.appendChild(heading);

    if (isEmptyTree(tree)) {
      const empty: HTMLElement = document.createElement('p');
      empty.className = 'mdv-related-empty';
      empty.textContent = strings.FolderPageEmptyMessage;
      this._contentElement.appendChild(empty);
    } else {
      this._contentElement.appendChild(
        renderTree(document, tree, {
          buildDocUrl: (docPath: string): string => this._buildDocUrl(docPath),
          linkAttribute: DOC_LINK_ATTRIBUTE,
          currentPath: path,
          linkFolders: true,
          restrictedFolders: library.getRestrictedFolders(),
          lockedLabel: strings.PermLockedShortLabel
        })
      );
    }

    this._renderSidePanels();

    // The chrome still wants the folder's own item, for its tags and its padlock.
    void this._loadItemState(path, token);
  }

  /**
   * Brings the heading a `#...` in the address refers to into view.
   *
   * The ids come from the document's own headings, so they can hold anything —
   * `#1-ビルド` is not a valid CSS selector. Walking the elements avoids having
   * to escape them.
   */
  private _scrollToAnchor(hash: string): void {
    if (!this._contentElement || !hash) {
      return;
    }

    let id: string = hash.charAt(0) === '#' ? hash.slice(1) : hash;
    try {
      id = decodeURIComponent(id);
    } catch {
      // Malformed percent-encoding: match on what was given.
    }

    if (id === '') {
      return;
    }

    const candidates: NodeListOf<Element> = this._contentElement.querySelectorAll('[id]');
    for (let i: number = 0; i < candidates.length; i++) {
      if (candidates[i].getAttribute('id') === id) {
        candidates[i].scrollIntoView();
        return;
      }
    }
  }

  private _createLinkContext(path: string, libraryUrl: string): IDocLinkContext {
    const provider: SharePointDocumentProvider = this._provider as SharePointDocumentProvider;

    return {
      currentPath: path,
      buildDocUrl: (docPath: string): string => this._buildDocUrl(docPath),
      buildFileUrl: (filePath: string): string => provider.getFileAbsoluteUrl(libraryUrl, filePath)
    };
  }

  private _buildDocUrl(path: string): string {
    return window.location.pathname + this._buildDocSearch(path);
  }

  /**
   * `?file=...`, built from scratch rather than edited.
   *
   * Only the chrome setting survives a navigation. Everything else in the
   * address — a tag search, a search request, a help page — describes what was
   * on screen a moment ago, and carrying `?q=...` into the next document left
   * an address that no longer said what it showed.
   */
  private _buildDocSearch(path: string): string {
    const search: string = setQueryParam('', QUERY_PARAMETER, path);
    const chrome: string = getQueryParam(window.location.search, CHROME_PARAMETER);

    return chrome === '' ? search : setQueryParam(search, CHROME_PARAMETER, chrome);
  }

  /** Absolute URL of one heading of the document on screen. */
  private _buildAnchorUrl(anchor: string): string {
    return window.location.origin + this._buildDocUrl(this._currentPath) + '#' + anchor;
  }

  /**
   * Where a tag chip points. With no page configured the search happens on this
   * page, which keeps the back button working and needs no setup; a configured
   * page (`search_tag.aspx`) gets the same web part with `?tag=` instead.
   */
  private _buildTagUrl(tag: string): string {
    const page: string = this._getTagSearchPage();
    const search: string = setQueryParam(
      page === '' ? removeQueryParam(window.location.search, QUERY_PARAMETER) : '',
      TAG_PARAMETER,
      tag
    );

    return (page === '' ? window.location.pathname : page) + search;
  }

  /** Configured search page as a URL, or '' when the search stays in place. */
  private _getTagSearchPage(): string {
    const configured: string = (this.properties.tagSearchPageUrl || '').replace(/\?.*$/, '');
    if (configured === '') {
      return '';
    }

    if (/^https?:\/\//i.test(configured) || configured.charAt(0) === '/') {
      return configured;
    }

    const web: string = this.context.pageContext.web.serverRelativeUrl.replace(/\/+$/, '');
    return web + '/SitePages/' + configured;
  }

  private _setPathLabels(path: string): void {
    if (this._footerPathElement) {
      this._footerPathElement.textContent = path;
    }
  }

  /**
   * The page shown when a document could not be read.
   *
   * "File not found" is only half true and sends people looking for a typo:
   * SharePoint answers the same way for a file that exists and is not theirs to
   * see, which is the likelier case now that pages can be restricted. The tree
   * of what *is* readable nearby is what turns a dead end into a way onwards.
   */
  private _showNotFound(path: string, token: number): void {
    if (!this._contentElement) {
      return;
    }

    this._contentElement.innerHTML = '';

    const status: HTMLElement = document.createElement('div');
    status.className = styles.status;

    const message: HTMLElement = document.createElement('p');
    message.className = 'mdv-notfound-message';
    message.textContent = strings.NotFoundMessage;
    status.appendChild(message);

    const reason: HTMLElement = document.createElement('p');
    reason.className = 'mdv-notfound-reason';
    reason.textContent = strings.NotFoundReasonMessage;
    status.appendChild(reason);

    const back: HTMLButtonElement = document.createElement('button');
    back.className = 'mdv-btn';
    back.setAttribute('type', 'button');
    back.textContent = strings.BackLabel;
    back.onclick = (): void => {
      // Whatever they were reading before, rather than a guess at a parent that
      // may be just as unreadable.
      window.history.back();
    };
    status.appendChild(back);

    const treeHost: HTMLElement = document.createElement('div');
    treeHost.className = 'mdv-notfound-tree';
    status.appendChild(treeHost);

    this._contentElement.appendChild(status);

    void this._renderNotFoundTree(path, treeHost, token);
  }

  /** The readable neighbourhood of a document that could not be opened. */
  private async _renderNotFoundTree(
    path: string,
    host: HTMLElement,
    token: number
  ): Promise<void> {
    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library || token !== this._loadToken) {
      return;
    }

    let pages: IRelatedPage[];
    try {
      pages = await library.getPages(await this._getConfiguredTagFields());
    } catch {
      return; // The message above stands on its own.
    }

    if (token !== this._loadToken) {
      return;
    }

    // Rooted at the library, so the folders leading to the missing file are
    // visible even when the folder itself is unreadable.
    const tree: ITreeFolder = buildTree(pages, '');
    if (isEmptyTree(tree)) {
      return;
    }

    const label: HTMLElement = document.createElement('div');
    label.className = 'mdv-help-label';
    label.textContent = strings.NotFoundTreeLabel;
    host.appendChild(label);

    host.appendChild(
      renderTree(document, tree, {
        buildDocUrl: (docPath: string): string => this._buildDocUrl(docPath),
        linkAttribute: DOC_LINK_ATTRIBUTE,
        currentPath: path,
        rootLabel: strings.RootLabel,
        linkFolders: true,
        restrictedFolders: library.getRestrictedFolders(),
        lockedLabel: strings.PermLockedShortLabel
      })
    );
  }

  /**
   * Publishes the height of the sticky chrome as a CSS variable.
   *
   * The side pane has to begin where that chrome ends, and the height depends
   * on how many tag columns are configured and whether a panel is open — so it
   * is measured rather than guessed. Every path that changes the header calls
   * this.
   */
  private _measureTopPanel(): void {
    if (!this._rootElement || !this._topPanelElement) {
      return;
    }

    this._rootElement.style.setProperty(
      '--mdv-top',
      this._topPanelElement.offsetHeight + 'px'
    );
  }

  private _showStatus(className: string, message: string): void {
    if (this._contentElement) {
      this._contentElement.innerHTML = `<div class="${className}">${escape(message)}</div>`;
    }
  }

  private _showError(message: string, detail: string | undefined): void {
    if (!this._contentElement) {
      return;
    }
    const detailHtml: string = detail
      ? `<span class="${styles.errorDetail}">${escape(detail)}</span>`
      : '';
    this._contentElement.innerHTML = `<div class="${styles.error}">${escape(message)}${detailHtml}</div>`;
  }

  // ------------------------------------------------------------- item state

  private async _ensureLibraryService(): Promise<SharePointLibraryService | undefined> {
    if (this._library) {
      return this._library;
    }
    if (!this._provider) {
      return undefined;
    }

    try {
      const libraryUrl: string = await this._provider.getLibraryServerRelativeUrl();
      this._library = new SharePointLibraryService(
        this.context.spHttpClient,
        this._provider.siteAbsoluteUrl,
        libraryUrl
      );
      return this._library;
    } catch {
      return undefined;
    }
  }

  /** Loads the list item behind the document, then the tags. */
  private async _loadItemState(path: string, token: number): Promise<void> {
    const hasTagFields: boolean = this._getTagFieldNames().length > 0;
    const needsMetadata: boolean =
      this._isEnabled('showEditButton') ||
      this._isEnabled('showOpenInAppButton') ||
      hasTagFields;

    if (!needsMetadata) {
      return;
    }

    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library || token !== this._loadToken) {
      return;
    }

    const fields: ITagFieldInfo[] = await this._getConfiguredTagFields();
    if (token !== this._loadToken) {
      return;
    }

    try {
      // A folder has a list item like any other, but it is not reachable
      // through the file endpoint.
      const metadata: IDocumentMetadata | undefined = isFolderPath(path)
        ? await library.getFolderMetadata(path, fields)
        : await library.getDocumentMetadata(path, fields);

      if (token !== this._loadToken) {
        return;
      }

      this._metadata = metadata;
      this._hasUniquePermissions = metadata ? metadata.hasUniquePermissions : false;
      this._renderToolbar();
      this._renderTagBar();
    } catch {
      return;
    }

  }

  // -------------------------------------------------------------- toolbar UI

  private _onRootClick = (event: MouseEvent): void => {
    if (!isPlainLeftClick(event) || !this._rootElement) {
      return;
    }

    const action: ToolbarAction | undefined = this._findAction(event.target as Node | null);
    if (action) {
      event.preventDefault();
      this._runAction(action);
      return;
    }

    const link: string = findDocLinkPath(event.target as Node | null, this._rootElement);
    // Folders are destinations too now, so a link to one has to survive this.
  const path: string = sanitizeDocPath(link, true);
    if (path !== '') {
      event.preventDefault();
      // `install.md#step-1` has to land on the step, not at the top.
      void this._navigate(path, true, splitLink(link).hash);
      return;
    }

    // Help links come before document links: they are inside the help page,
    // which has no library document to resolve them against.
    const helpKey: string | undefined = findHelpLink(
      event.target as Node | null,
      this._rootElement
    );
    if (helpKey !== undefined) {
      event.preventDefault();
      this._openHelp(helpKey, true);
      return;
    }

    const code: string = findCodeToCopy(event.target as Node | null, this._rootElement);
    if (code !== '') {
      event.preventDefault();
      void copyToClipboard(code);
      this._showToast(strings.CodeCopiedMessage);
      return;
    }

    // Heading controls come last: a link inside a heading has to win.
    const anchor: string = findHeadingAnchor(event.target as Node | null, this._rootElement);
    if (anchor !== '') {
      event.preventDefault();
      void copyToClipboard(this._buildAnchorUrl(anchor));
      this._showToast(strings.LinkCopiedMessage);
      return;
    }

    if (this._contentElement && toggleHeadingAt(event.target as Node | null, this._contentElement)) {
      event.preventDefault();
    }
  };

  private _findAction(start: Node | null): ToolbarAction | undefined {
    let node: Node | null = start;

    while (node && node !== this._rootElement) {
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

  private _runAction(action: ToolbarAction): void {
    // The menu works without a list item: it is how an empty library gets its
    // first page.
    if (action === 'menu') {
      if (this._navPanelOpen) {
        this._closeNavPanel();
      } else {
        this._onMenuClick(dirname(this._currentPath));
      }
      return;
    }

    if (action === 'home') {
      void this._navigate(sanitizeDocPath(this.properties.startFile || ''), true);
      return;
    }

    if (!this._metadata) {
      return;
    }

    if (action === 'path') {
      if (this._renamePanelOpen) {
        this._closeRenamePanel();
      } else {
        this._openRenamePanel();
      }
      return;
    }

    if (action === 'edit') {
      window.open(this._metadata.editUrl, '_blank', 'noopener,noreferrer');
      return;
    }

    if (action === 'outline') {
      this._onOutlineClick();
      return;
    }

    if (action === 'permissions') {
      if (this._permPanelOpen) {
        this._closePermPanel();
      } else {
        void this._openPermPanel();
      }
      return;
    }

    if (action === 'open-app') {
      this._openInApp(this._metadata);
      return;
    }

  }

  // ------------------------------------------------------------- open in app

  /**
   * Hands the document to whatever application is registered for the scheme.
   *
   * Just the URL — where the file sits on disk is the application's business,
   * not this page's.
   */
  private _openInApp(metadata: IDocumentMetadata): void {
    const scheme: string = (this.properties.appProtocol || DEFAULT_APP_SCHEME).replace(/:.*$/, '');

    // An unregistered scheme fails silently in every browser, so the only way
    // to tell is to check whether we still have focus a moment later. Saying so
    // beats a button that appears to do nothing.
    window.setTimeout((): void => {
      if (document.hasFocus()) {
        this._showNotice(strings.OpenInAppFailedMessage.replace('{0}', scheme));
      }
    }, PROTOCOL_TIMEOUT_MS);

    try {
      window.location.href = buildProtocolUrl(scheme, metadata.fileUrl);
    } catch {
      this._showNotice(strings.OpenInAppFailedMessage.replace('{0}', scheme));
    }
  }

  private _clearNotice(): void {
    if (this._noticeElement) {
      this._noticeElement.innerHTML = '';
    }
  }

  /**
   * A confirmation that fades out on its own, bottom right.
   *
   * Anything the reader has to act on — a failure, a path to copy — goes to
   * {@link _showNotice} instead and stays until dismissed.
   */
  private _showToast(message: string): void {
    if (this._toastElement) {
      showToast({ container: this._toastElement, message: message });
    }
  }

  // ------------------------------------------------------------- tag editing

  /**
   * Opens the editor for one column.
   *
   * One column at a time, because that is what the reader asked for by clicking
   * its name: a panel holding all three made it far too easy to save a change
   * to a column nobody meant to touch.
   */
  private async _openTagPanel(internalName: string): Promise<void> {
    if (!this._tagPanelElement || !this._metadata) {
      return;
    }

    let group: ITagGroup | undefined;
    for (let i: number = 0; i < this._metadata.tagGroups.length; i++) {
      if (this._metadata.tagGroups[i].field.internalName === internalName) {
        group = this._metadata.tagGroups[i];
        break;
      }
    }

    if (!group) {
      this._showNotice(strings.TagFieldMissingMessage);
      return;
    }

    this._tagPanelOpen = true;
    this._tagPanelField = internalName;
    const itemId: number = this._metadata.itemId;

    this._tagEditor = renderTagEditor({
      container: this._tagPanelElement,
      groups: [group],
      strings: {
        title: strings.TagPanelTitle,
        save: strings.SaveLabel,
        cancel: strings.CancelLabel,
        textHint: strings.TagTextHint,
        noChoices: strings.TagNoChoices,
        addLabel: strings.TagAddLabel,
        addPlaceholder: strings.TagAddPlaceholder,
        addDisabledPlaceholder: strings.TagAddDisabledPlaceholder,
        addHint: strings.TagAddHint,
        addDisabledHint: strings.TagAddDisabledHint
      },
      onSave: (edited: ITagGroup[]): void => {
        void this._saveTags(itemId, edited);
      },
      onCancel: (): void => this._closeTagPanel()
    });
  }

  private _closeTagPanel(): void {
    this._tagPanelOpen = false;
    this._tagPanelField = '';
    this._tagEditor = undefined;
    if (this._tagPanelElement) {
      this._tagPanelElement.innerHTML = '';
    }
  }

  /**
   * The configured columns, resolved against the library's column list.
   * A column that no longer exists is dropped rather than failing the load.
   */
  private async _getConfiguredTagFields(): Promise<ITagFieldInfo[]> {
    const names: string[] = this._getTagFieldNames();
    if (names.length === 0) {
      return [];
    }

    if (this._tagFields.length === 0) {
      try {
        await this._loadTagFields();
      } catch {
        return [];
      }
    }

    const fields: ITagFieldInfo[] = [];
    for (let i: number = 0; i < names.length; i++) {
      for (let f: number = 0; f < this._tagFields.length; f++) {
        if (this._tagFields[f].internalName === names[i]) {
          fields.push(this._tagFields[f]);
          break;
        }
      }
    }

    return fields;
  }

  private async _saveTags(itemId: number, groups: ITagGroup[]): Promise<void> {
    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library) {
      return;
    }

    try {
      await library.setTags(itemId, groups);
    } catch (error) {
      const message: string = strings.TagSaveFailedMessage + ' ' + describeError(error);

      // Inside the panel, next to the button that was just pressed, so the
      // edit that failed is still on screen to correct or abandon.
      if (this._tagEditor) {
        this._tagEditor.setError(message);
      } else {
        this._showNotice(message);
      }
      return;
    }

    // No confirmation: a failure keeps the panel open with the message in it, so
    // the panel closing is already the answer to "did that work?".
    this._closeTagPanel();

    if (this._metadata) {
      // `groups` only covers the column that was open, so the rest have to be
      // carried over rather than overwritten.
      const merged: ITagGroup[] = mergeTagGroups(this._metadata.tagGroups, groups);

      this._metadata = {
        itemId: this._metadata.itemId,
        editUrl: this._metadata.editUrl,
        fileUrl: this._metadata.fileUrl,
        tagGroups: merged,
        tags: flattenTagGroups(merged),
        hasUniquePermissions: this._metadata.hasUniquePermissions
      };
      this._renderTagBar();
    }

  }

  // ------------------------------------------------------- menu and file work

  /** Opens the popup for `folder`, loading what that folder holds. */
  private async _openNavPanel(folder: string): Promise<void> {
    const host: HTMLElement | undefined =
      this._pinnedPanel === 'tree' ? this._sidePaneElement : this._navElement;

    if (!host) {
      return;
    }

    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library) {
      this._showNotice(strings.LibraryNotFoundMessage);
      return;
    }

    let listing: IFolderListing;
    try {
      listing = await library.listFolder(folder);
    } catch (error) {
      this._showNotice(strings.ErrorMessage + ' ' + describeError(error));
      return;
    }

    this._navPanelOpen = this._pinnedPanel !== 'tree';

    renderNavPanel({
      container: host,
      ancestors: buildAncestors(folder),
      folders: listing.folders,
      files: listing.files,
      currentPath: this._currentPath,
      strings: {
        newPage: strings.NewPageTitle,
        newPageName: strings.NewPageNamePlaceholder,
        format: strings.NewPageFormatLabel,
        openWith: strings.NewPageOpenWithLabel,
        openInSharePoint: strings.NewPageOpenInSharePoint,
        openInApp: strings.NewPageOpenInApp,
        create: strings.CreateLabel,
        cancel: strings.CancelLabel,
        location: strings.LocationTitle,
        root: strings.RootLabel,
        empty: strings.FolderEmptyMessage,
        close: strings.CloseLabel,
        search: strings.SearchLabel,
        help: strings.HelpLabel,
        chromeToggle: this._isFullScreen()
          ? strings.ChromeSharePointLabel
          : strings.ChromeFullScreenLabel
      },
      buildDocUrl: (docPath: string): string => this._buildDocUrl(docPath),
      onOpen: (docPath: string): void => {
        this._closeNavPanel();
        void this._navigate(docPath, true);
      },
      onFolder: (folderPath: string): void => {
        void this._openNavPanel(folderPath);
      },
      onCreate: (request: INewPageRequest): void => {
        void this._createPage(folder, request);
      },
      onClose: (): void => this._closeNavPanel(),
      onSearch: (): void => {
        this._closeNavPanel();
        this._openSearch({ term: '', mode: 'name', scope: folder });
      },
      onHelp: (): void => {
        this._closeNavPanel();
        this._openHelp('', true);
      },
      onChromeToggle: (): void => this._toggleChrome()
    });
  }

  /** True when the page is being shown without the SharePoint chrome. */
  private _isFullScreen(): boolean {
    return getQueryParam(window.location.search, CHROME_PARAMETER) !== '';
  }

  /**
   * Switches between the site chrome and the full screen view.
   *
   * `env` is read by SharePoint's page framework, not by this web part, so the
   * only way to change it is to load the page again with it added or removed.
   */
  private _toggleChrome(): void {
    const search: string = this._isFullScreen()
      ? removeQueryParam(window.location.search, CHROME_PARAMETER)
      : setQueryParam(window.location.search, CHROME_PARAMETER, FULL_SCREEN_CHROME);

    window.location.href = window.location.pathname + search + window.location.hash;
  }

  /**
   * Shows the help page in place of the document.
   *
   * `pushHistory` is false when arriving through the address bar or the back
   * button, where the address already says what is on screen.
   */
  private _openHelp(docKey: string, pushHistory: boolean): void {
    if (!this._contentElement) {
      return;
    }

    this._clearView();
    this._helpOpen = true;

    if (pushHistory) {
      window.history.pushState(
        { help: docKey },
        '',
        window.location.pathname + this._buildHelpSearch(docKey)
      );
    }

    renderHelpPage({
      container: this._contentElement,
      docKey: docKey,
      docs: BUNDLED_DOCS,
      setup: {
        viewerPageUrl: window.location.pathname,
        // Only when it is not the library the command would pick by itself.
        libraryRootUrl: this._getConfiguredLibraryRootUrl(),
        chromeMode: getQueryParam(window.location.search, CHROME_PARAMETER) || FULL_SCREEN_CHROME,
        extensionsListUrl: this._buildExtensionsListUrl()
      },
      strings: {
        title: strings.HelpTitle,
        documents: strings.HelpDocumentsTitle,
        setupTitle: strings.HelpSetupTitle,
        setupIntro: strings.HelpCommandSetIntro,
        componentId: strings.HelpComponentId,
        location: strings.HelpLocation,
        listTemplate: strings.HelpListTemplate,
        componentProperties: strings.HelpComponentProperties,
        extensionsList: strings.HelpExtensionsList,
        copy: strings.CopyLabel,
        copied: strings.CopiedLabel,
        backToIndex: strings.HelpBackLabel
      },
      buildHelpUrl: (key: string): string =>
        window.location.pathname + this._buildHelpSearch(key),
      renderMarkdown: (markdown: string, into: HTMLElement): void => {
        // The same pipeline as a library document, so a heading, a table or a
        // code block looks the same wherever the reader met it.
        renderDocument({
          source: markdown,
          container: into,
          sanitize: sanitizeHtml,
          format: 'markdown',
          context: {
            currentPath: '',
            // The manual's own links are rewritten to `?help=` at build time;
            // anything left pointing into a library would be a broken guess.
            buildDocUrl: (): string => '',
            buildFileUrl: (): string => ''
          },
          headings: {
            copyLabel: strings.HeadingCopyLabel,
            toggleLabel: strings.HeadingToggleLabel
          },
          codeBlocks: { copyLabel: strings.CodeCopyLabel }
        });
      },
      onCopy: (value: string): void => {
        void copyToClipboard(value);
      }
    });
  }

  /** `?help=<key>`, keeping the chrome setting the reader arrived with. */
  private _buildHelpSearch(docKey: string): string {
    const kept: string = setQueryParam('', HELP_PARAMETER, docKey);
    const chrome: string = getQueryParam(window.location.search, CHROME_PARAMETER);

    return chrome === '' ? kept : setQueryParam(kept, CHROME_PARAMETER, chrome);
  }

  /**
   * The library root the command set has to be told about, or '' when it can
   * work it out itself.
   *
   * The command defaults to the library the reader invoked it from, which is
   * right unless this web part is pointed at a folder inside that library.
   */
  private _getConfiguredLibraryRootUrl(): string {
    const name: string = (this.properties.libraryName || '').replace(/^\/+|\/+$/g, '');
    if (name.indexOf('/') < 0) {
      return '';
    }

    const web: string = this.context.pageContext.web.serverRelativeUrl.replace(/\/+$/, '');
    return web + '/' + name;
  }

  /** Address of the tenant wide extensions list, derived from this tenant. */
  private _buildExtensionsListUrl(): string {
    const match: RegExpExecArray | null = /^(https?:\/\/[^/]+)/i.exec(
      this.context.pageContext.web.absoluteUrl
    );

    return match ? match[1] + '/sites/appcatalog/Lists/TenantWideExtensions' : '';
  }

  private _closeNavPanel(): void {
    this._navPanelOpen = false;
    if (this._navElement) {
      this._navElement.innerHTML = '';
    }
  }

  /**
   * Creates a document from the template, then hands it to the editor the
   * writer asked for.
   *
   * SharePoint's editor replaces the page rather than opening a tab: a tab
   * opened this long after the click is what pop-up blockers stop, and the
   * reader can come back with the browser's back button. The desktop app is a
   * URL scheme, so the viewer stays where it is and shows the new page.
   */
  private async _createPage(folder: string, request: INewPageRequest): Promise<void> {
    const fileName: string = withExtension(request.name, request.format);
    if (fileName === '') {
      this._showNotice(strings.InvalidNameMessage);
      return;
    }

    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library) {
      return;
    }

    let path: string;
    try {
      path = await library.createFile(
        folder,
        fileName,
        newDocumentTemplate(fileName, request.format)
      );
    } catch (error) {
      this._showNotice(strings.CreateFailedMessage + ' ' + describeError(error));
      return;
    }

    this._closeNavPanel();
    library.invalidatePages();

    const fields: ITagFieldInfo[] = await this._getConfiguredTagFields();
    let metadata: IDocumentMetadata | undefined;
    try {
      metadata = await library.getDocumentMetadata(path, fields);
    } catch {
      metadata = undefined;
    }

    if (!metadata) {
      // The file exists; only the follow-up could not be prepared.
      await this._navigate(path, true);
      this._showNotice(strings.NewPageCreatedMessage);
      return;
    }

    if (request.editor === 'app') {
      await this._navigate(path, true);
      this._openInApp(metadata);
      return;
    }

    window.location.href = metadata.editUrl;
  }

  private _openRenamePanel(): void {
    if (!this._renamePanelElement) {
      return;
    }

    this._renamePanelOpen = true;
    // Fills the suggestions in as they arrive; the panel is usable meanwhile.
    void this._loadFolders();

    renderRenamePanel({
      container: this._renamePanelElement,
      path: this._currentPath,
      folders: this._folders,
      strings: {
        title: strings.RenameTitle,
        folder: strings.RenameFolderLabel,
        name: strings.RenameNameLabel,
        hint: strings.RenameHint,
        save: strings.SaveLabel,
        cancel: strings.CancelLabel
      },
      onSave: (path: string): void => {
        void this._rename(path);
      },
      onCancel: (): void => this._closeRenamePanel()
    });
  }

  /**
   * Reads the library's folders once, for the rename panel's suggestions.
   * Re-renders the panel when they arrive so it does not have to wait for them.
   */
  private async _loadFolders(): Promise<void> {
    if (this._folders.length > 0) {
      return;
    }

    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library) {
      return;
    }

    try {
      this._folders = await library.listAllFolders();
    } catch {
      return; // Typing the folder still works.
    }

    if (this._renamePanelOpen && this._folders.length > 0) {
      this._openRenamePanel();
    }
  }

  private _closeRenamePanel(): void {
    this._renamePanelOpen = false;
    if (this._renamePanelElement) {
      this._renamePanelElement.innerHTML = '';
    }
  }

  private async _rename(rawPath: string): Promise<void> {
    const path: string = normalizePath(rawPath);

    if (path === '' || !isDocumentPath(path)) {
      this._showNotice(strings.InvalidNameMessage);
      return;
    }

    if (path.toLowerCase() === this._currentPath.toLowerCase()) {
      this._closeRenamePanel();
      return;
    }

    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library) {
      return;
    }

    try {
      await library.moveFile(this._currentPath, path);
    } catch (error) {
      this._showNotice(strings.RenameFailedMessage + ' ' + describeError(error));
      return;
    }

    this._closeRenamePanel();
    await this._navigate(path, true);
  }

  // ------------------------------------------------------------- tag search

  /** Replaces the document with the list of every page carrying `tag`. */
  private async _searchTag(tag: string, pushHistory: boolean): Promise<void> {
    if (!this._contentElement) {
      return;
    }

    const page: string = this._getTagSearchPage();
    if (page !== '') {
      // A page of its own was configured, so let the browser go there.
      window.location.href = this._buildTagUrl(tag);
      return;
    }

    this._closePanels();
    this._clearNotice();
    if (pushHistory) {
      window.history.pushState({ tag: tag }, '', this._buildTagUrl(tag));
    }

    this._loadToken += 1;
    const token: number = this._loadToken;
    this._showStatus(styles.status, strings.LoadingMessage);

    const library: SharePointLibraryService | undefined = await this._ensureLibraryService();
    if (!library || token !== this._loadToken) {
      return;
    }

    let pages: IRelatedPage[];
    try {
      pages = await library.getPages(await this._getConfiguredTagFields());
    } catch (error) {
      this._showError(strings.ErrorMessage, describeError(error));
      return;
    }

    if (token !== this._loadToken || !this._contentElement) {
      return;
    }

    renderTagSearch({
      container: this._contentElement,
      tag: tag,
      pages: filterPagesByTag(pages, tag),
      title: strings.TagSearchTitle,
      emptyMessage: strings.TagSearchEmpty,
      linkAttribute: DOC_LINK_ATTRIBUTE,
      buildDocUrl: (docPath: string): string => this._buildDocUrl(docPath)
    });
  }

  private _closePanels(): void {
    this._closePermPanel();
    this._closeTagPanel();
    this._closeRenamePanel();
    this._closeNavPanel();
  }

  /**
   * Takes the whole document off screen — tags, panels, notices and body.
   *
   * The search and help pages are not the document with something added, they
   * are a different thing entirely; leaving the tag bar of whatever was open
   * before sitting above them made them look like part of that page.
   */
  private _clearView(): void {
    this._closePanels();
    this._clearNotice();
    this._clearTagBar();
    this._setPathLabels('');

    this._currentPath = '';
    this._metadata = undefined;
    this._hasUniquePermissions = undefined;

    if (this._contentElement) {
      this._contentElement.innerHTML = '';
    }
  }

  private _clearTagBar(): void {
    if (this._tagBarElement) {
      this._tagBarElement.innerHTML = '';
    }
  }

  private _showNotice(message: string): void {
    if (!this._noticeElement) {
      return;
    }

    renderNotice({
      container: this._noticeElement,
      message: message,
      closeLabel: strings.CloseLabel,
      onClose: (): void => this._clearNotice()
    });
  }

  // ---------------------------------------------------------------- history

  private _readPathFromUrl(): string {
    const fromQuery: string = sanitizeDocPath(
      getQueryParam(window.location.search, QUERY_PARAMETER),
      true
    );
    return fromQuery !== '' ? fromQuery : sanitizeDocPath(this.properties.startFile || '');
  }

  /**
   * Double clicks, for the controls that have a second meaning.
   *
   * Separate from the click handler because a double click also fires two
   * clicks; the click handler defers, and this cancels what it deferred.
   */
  private _onRootDoubleClick = (event: MouseEvent): void => {
    if (!this._rootElement) {
      return;
    }

    const action: ToolbarAction | undefined = this._findAction(event.target as Node | null);

    if (action === 'outline') {
      event.preventDefault();
      this._onOutlineDoubleClick();
      return;
    }

    if (action === 'menu') {
      event.preventDefault();
      this._onMenuDoubleClick(dirname(this._currentPath));
    }
  };

  private _onPopState = (): void => {
    // Same order as the initial render, so the back button lands on whatever
    // the address says rather than on whatever happens to be in memory.
    if (getQueryParam(window.location.search, HELP_PARAMETER) !== '') {
      this._openHelp(getQueryParam(window.location.search, HELP_PARAMETER), false);
      return;
    }

    if (getQueryParam(window.location.search, SEARCH_PARAMETER) !== '') {
      this._openSearch({
        term: getQueryParam(window.location.search, SEARCH_PARAMETER),
        mode:
          getQueryParam(window.location.search, SEARCH_MODE_PARAMETER) === 'content'
            ? 'content'
            : 'name',
        scope: getQueryParam(window.location.search, SEARCH_SCOPE_PARAMETER),
        restrictedOnly:
          getQueryParam(window.location.search, SEARCH_RESTRICTED_PARAMETER) === '1'
      });
      return;
    }

    const tag: string = getQueryParam(window.location.search, TAG_PARAMETER);
    if (tag !== '') {
      void this._searchTag(tag, false);
      return;
    }

    this._helpOpen = false;
    this._searchOpen = false;
    void this._navigate(this._readPathFromUrl(), false);
  };
}

/**
 * Recovers the server relative library root from a file URL and the library
 * relative path of that same file.
 */
export function extractLibraryUrl(fileAbsoluteUrl: string, path: string): string {
  const match: RegExpExecArray | null = /^https?:\/\/[^/]+(\/.*)$/i.exec(fileAbsoluteUrl);
  if (!match) {
    return '';
  }

  const serverRelative: string = decodeURI(match[1]);
  const suffix: string = '/' + path;
  const index: number = serverRelative.toLowerCase().lastIndexOf(suffix.toLowerCase());

  return index < 0 ? serverRelative : serverRelative.slice(0, index);
}

async function copyToClipboard(value: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(value);
  } catch {
    // Clipboard access can be denied; the path is on screen either way.
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function describeError(error: any): string {
  if (!error) {
    return '';
  }
  return typeof error === 'string' ? error : error.message || String(error);
}
