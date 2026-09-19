import { BaseDialog, type IDialogConfiguration } from '@microsoft/sp-dialog';

import styles from './NoticeDialog.module.scss';
import { renderNotice } from '../../webparts/spWiki/core/chrome';

export interface INoticeDialogOptions {
  message: string;
  /** Monospace detail line, typically a file path. */
  detail?: string;
  /** Omit to hide the copy button. */
  copyLabel?: string;
  closeLabel: string;
}

/**
 * One-line notice shown from a list view, where there is no web part surface to
 * draw into. Used for the "open in desktop app" fallback, so the command and
 * the web part button behave the same when the URL scheme is not registered.
 */
export default class NoticeDialog extends BaseDialog {
  private readonly _options: INoticeDialogOptions;

  public constructor(options: INoticeDialogOptions) {
    super({ isBlocking: false });
    this._options = options;
  }

  public render(): void {
    const host: HTMLDivElement = document.createElement('div');
    host.className = styles.notice;
    this.domElement.appendChild(host);

    renderNotice({
      container: host,
      message: this._options.message,
      detail: this._options.detail,
      copyLabel: this._options.copyLabel,
      onCopy: (): void => {
        if (this._options.detail) {
          void copyToClipboard(this._options.detail);
        }
      },
      closeLabel: this._options.closeLabel,
      onClose: (): void => {
        this.close().catch((): void => undefined);
      }
    });
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: false };
  }
}

async function copyToClipboard(value: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(value);
  } catch {
    // Clipboard access can be denied; the path is on screen either way.
  }
}
