import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Swal from 'sweetalert2';

export interface ILogItem {
  Id?: number;
  Title: string;
  Page: string;
  ItemId: string;
  Module: string;
  Severity: 'High' | 'Medium' | 'Low';
  Error: string;
  ErrorId?: string;
  Created?: string;
  Author?: { Title: string };
}

export class LoggerService {
  private static _sp: SPFI;
  private static _logListTitle: string = "";
  private static _currentListName: string = "System";
  private static _showAlerts: boolean = true;
  public static enabled: boolean = false;

  public static init(context: WebPartContext, sourceList: string, logList: string, enabled: boolean, showAlerts: boolean): void {
    this._sp = spfi().using(SPFx(context));
    this._currentListName = sourceList || "System";
    this._logListTitle = logList;
    this.enabled = enabled && !!logList;
    this._showAlerts = showAlerts;
  }

  private static _generateGuid(): string {
    return crypto.randomUUID(); 
  }

  private static async _copyToClipboard(text: string): Promise<void> {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
      } else {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.style.position = "fixed";
        textArea.style.opacity = "0";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
      }
    } catch (err:any) {
      console.warn('[LoggerService] Clipboard copy failed', err);
    }
  }

  public static async log(
    module: string,
    severity: 'High' | 'Medium' | 'Low',
    itemId: string = 'N/A',
    errorMsg: string
  ): Promise<void> {
    if (!this.enabled || !this._logListTitle || !this._sp) return;

    const fullUrl = window.location.href;
    const pathName = window.location.pathname;
    
    const listPart = this._currentListName.replace(/[^a-zA-Z0-9]/g, '').substring(0, 5);
    const guid = this._generateGuid().substring(0, 8);
    const errorId = `${listPart}_${guid}`.toUpperCase();

    if (this._showAlerts && (severity === 'High' || severity === 'Medium')) {
      void Swal.fire({
        icon: 'error',
        title: 'Unexpected Error',
        html: `
          <div style="text-align: left; font-size: 14px;">
            <p>Something went wrong. Please contact the administrator.</p>
            <p><b>Reference ID:</b></p>
            <div style="background: #f3f3f3; padding: 8px; border: 1px dashed #ccc; word-break: break-all;">
              ${errorId}
            </div>
          </div>
        `,
        showCancelButton: true,
        cancelButtonText: 'Close',
        confirmButtonText: 'Copy ID',
        confirmButtonColor: '#0078d4',
        preConfirm: () => { void this._copyToClipboard(errorId); }
      });
    }

    try {
      await this._sp.web.lists.getByTitle(this._logListTitle).items.add({
        Title: this._currentListName,
        Page: fullUrl,
        ItemId: itemId,
        Module: module,
        Severity: severity,
        Error: errorMsg,
        ErrorId: errorId
      });
    } catch (e:any) {
      console.error('[LoggerService] CRITICAL: List write failed.', e);
    }
  }

  public static async getLogs(sourceList: string): Promise<ILogItem[]> {
    if (!this._logListTitle || !this._sp) return [];
    try {
      return await this._sp.web.lists.getByTitle(this._logListTitle).items
        .filter(`Title eq '${sourceList}'`)
        .select('Id', 'Title', 'Page', 'ItemId', 'Module', 'Severity', 'Error', 'ErrorId', 'Created', 'Author/Title')
        .expand('Author')
        .orderBy('Created', false)
        .top(5000)();
    } catch (e:any) {
      console.error("Fetch failed", e);
      return [];
    }
  }
}