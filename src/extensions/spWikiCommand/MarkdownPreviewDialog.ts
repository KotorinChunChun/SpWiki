import { BaseDialog, type IDialogConfiguration } from '@microsoft/sp-dialog';

import * as strings from 'SpWikiCommandSetStrings';

import viewerStyles from '../../styles/viewer.module.scss';
import dialogStyles from './MarkdownPreviewDialog.module.scss';

import {
  ACTION_ATTRIBUTE,
  ToolbarAction,
  renderToolbar
} from '../../webparts/spWiki/core/chrome';
import { findDocLinkPath, isPlainLeftClick } from '../../webparts/spWiki/core/domRewrite';
import { findHeadingAnchor, toggleHeadingAt } from '../../webparts/spWiki/core/headings';
import { sanitizeDocPath } from '../../webparts/spWiki/core/pathUtils';
import { renderDocument } from '../../webparts/spWiki/core/renderer';
import { sanitizeHtml } from '../../webparts/spWiki/core/sanitize';
import { IDocLinkContext, IDocumentResult } from '../../webparts/spWiki/core/types';
import { SharePointDocumentProvider } from '../../webparts/spWiki/services/SharePointDocumentProvider';

export interface IMarkdownPreviewDialogOptions {
  /** Library relative path of the document to show first. */
  path: string;
  /** Reads the document text; shared with the web part. */
  provider: SharePointDocumentProvider;
  /** Server relative library root, used to resolve images and other files. */
  libraryServerRelativeUrl: string;
  /** URL of the full page viewer for a given document, or '' when unset. */
  buildFullPageUrl(path: string): string;
}

/**
 * Shows a Markdown or HTML document in a modal over the document library, so a
 * library can be browsed without a page at all.
 *
 * Relative links keep working inside the dialog: clicking one loads the target
 * document in place, exactly as it does in the web part, because both go
 * through the same core modules.
 */
export default class MarkdownPreviewDialog extends BaseDialog {
  private readonly _options: IMarkdownPreviewDialogOptions;
  private _currentPath: string;
  private _contentElement: HTMLElement | undefined;
  private _toolbarElement: HTMLElement | undefined;
  private _loadToken: number = 0;

  public constructor(options: IMarkdownPreviewDialogOptions) {
    super({ isBlocking: false });
    this._options = options;
    this._currentPath = options.path;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${dialogStyles.dialog}">
        <div class="${viewerStyles.spWiki}">
          <div class="${viewerStyles.header}"></div>
        </div>
        <div class="${dialogStyles.body}">
          <div class="${viewerStyles.spWiki}">
            <div class="${viewerStyles.content}"></div>
          </div>
        </div>
      </div>`;

    this._toolbarElement = this.domElement.querySelector('.' + viewerStyles.header) as HTMLElement;
    this._contentElement = this.domElement.querySelector('.' + viewerStyles.content) as HTMLElement;

    this.domElement.addEventListener('click', this._onClick);

    void this._load(this._currentPath);
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: false };
  }

  protected onAfterClose(): void {
    this.domElement.removeEventListener('click', this._onClick);
    super.onAfterClose();
  }

  // ---------------------------------------------------------------- internals

  private _renderToolbar(): void {
    if (!this._toolbarElement) {
      return;
    }

    renderToolbar({
      container: this._toolbarElement,
      path: this._currentPath,
      buttons: [
        {
          action: 'full-page',
          label: strings.FullPageLabel,
          disabled: this._options.buildFullPageUrl(this._currentPath) === ''
        },
        { action: 'close', label: strings.CloseLabel }
      ]
    });
  }

  private async _load(path: string): Promise<void> {
    if (!this._contentElement) {
      return;
    }

    this._currentPath = path;
    this._renderToolbar();

    this._loadToken += 1;
    const token: number = this._loadToken;
    this._showStatus(strings.LoadingMessage);

    const result: IDocumentResult = await this._options.provider.getDocument(path);
    if (token !== this._loadToken || !this._contentElement) {
      return;
    }

    if (result.kind === 'notfound') {
      this._showStatus(strings.NotFoundMessage);
      return;
    }

    if (result.kind === 'error') {
      this._showStatus(strings.ErrorMessage + ' ' + (result.message || ''));
      return;
    }

    renderDocument({
      source: result.text || '',
      container: this._contentElement,
      sanitize: sanitizeHtml,
      context: this._createLinkContext(path),
      headings: {
        copyLabel: strings.HeadingCopyLabel,
        toggleLabel: strings.HeadingToggleLabel
      }
    });

    this._contentElement.scrollTop = 0;
  }

  private _createLinkContext(path: string): IDocLinkContext {
    return {
      currentPath: path,
      // A real href so ctrl-click and "copy link" reach the full page viewer,
      // while a plain click is intercepted and handled inside the dialog.
      buildDocUrl: (docPath: string): string => this._options.buildFullPageUrl(docPath),
      buildFileUrl: (filePath: string): string =>
        this._options.provider.getFileAbsoluteUrl(this._options.libraryServerRelativeUrl, filePath)
    };
  }

  private _showStatus(message: string): void {
    if (this._contentElement) {
      this._contentElement.innerHTML = '';
      const status: HTMLElement = document.createElement('div');
      status.className = viewerStyles.status;
      status.textContent = message;
      this._contentElement.appendChild(status);
    }
  }

  private _onClick = (event: MouseEvent): void => {
    if (!isPlainLeftClick(event)) {
      return;
    }

    const action: ToolbarAction | undefined = this._findAction(event.target as Node | null);
    if (action === 'close') {
      event.preventDefault();
      this.close().catch((): void => undefined);
      return;
    }

    if (action === 'full-page') {
      const url: string = this._options.buildFullPageUrl(this._currentPath);
      event.preventDefault();
      if (url !== '') {
        window.location.href = url;
      }
      return;
    }

    const path: string = sanitizeDocPath(findDocLinkPath(event.target as Node | null, this.domElement));
    if (path !== '') {
      event.preventDefault();
      void this._load(path);
      return;
    }

    // The dialog has no address of its own to copy, so the heading link points
    // at the full page viewer.
    const anchor: string = findHeadingAnchor(event.target as Node | null, this.domElement);
    if (anchor !== '') {
      event.preventDefault();
      const page: string = this._options.buildFullPageUrl(this._currentPath);
      if (page !== '') {
        void this._copyToClipboard(window.location.origin + page + '#' + anchor);
      }
      return;
    }

    if (this._contentElement && toggleHeadingAt(event.target as Node | null, this._contentElement)) {
      event.preventDefault();
    }
  };

  private async _copyToClipboard(value: string): Promise<void> {
    try {
      await navigator.clipboard.writeText(value);
    } catch {
      // Clipboard access can be denied; nothing else depends on it.
    }
  }

  private _findAction(start: Node | null): ToolbarAction | undefined {
    let node: Node | null = start;

    while (node && node !== this.domElement) {
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
}
