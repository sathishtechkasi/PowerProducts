import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from '@microsoft/sp-webpart-base';
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
    // --- GLOBAL STATE ---
    private static _logListTitle: string = "";
    private static _sp: SPFI; // Replaces the generic _context
    private static _currentListName: string = "System";
    public static enabled: boolean = false;
    private static _showAlerts: boolean = true;

    /**
     * INITIALIZE SERVICE
     * Called from WebPart.ts with dynamic user configurations 
     */
    public static init(context: WebPartContext, sourceList: string, logList: string, enabled: boolean, showAlerts: boolean): void {
        // Initialize the PnP v4 Factory using the SPFx context
        this._sp = spfi().using(SPFx(context));
        this._currentListName = sourceList || "System";
        this._logListTitle = logList; 
        this.enabled = enabled && !!logList; 
        this._showAlerts = showAlerts;
    }

    private static generateGuid(): string {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    private static fallbackCopy(text: string) {
        const textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.style.position = "fixed";
        textArea.style.top = "0";
        textArea.style.left = "0";
        textArea.style.width = "2em";
        textArea.style.height = "2em";
        textArea.style.padding = "0";
        textArea.style.border = "none";
        textArea.style.outline = "none";
        textArea.style.boxShadow = "none";
        textArea.style.background = "transparent";
        textArea.style.opacity = "0.01";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        try {
            const successful = document.execCommand('copy');
            if (!successful) {
                console.warn('[LoggerService] Fallback copy command failed.');
                prompt("Copy this Error ID manually:", text);
            }
        } catch (err) {
            console.error('[LoggerService] Fallback copy exception', err);
            prompt("Copy this Error ID manually:", text);
        }
        document.body.removeChild(textArea);
    }

    /**
     * FULL UPDATED LOG METHOD
     */
    public static async log(
        page: string,
        module: string,
        severity: 'High' | 'Medium' | 'Low',
        itemId: string = 'N/A',
        errorMsg: string
    ): Promise<void> {
        console.log(page, module, severity, itemId, errorMsg);
        if (!this.enabled || !this._logListTitle || !this._sp) {
            console.warn('[LoggerService] List not created or SPFI is NULL. Call init() first. Ignoring error:', errorMsg);
            return;
        }
        
        const fullUrl = page || window.location.href;
        const pathName = window.location.pathname;
        const pageName = pathName.substring(pathName.lastIndexOf('/') + 1) || 'Home';
        
        const listPart = (this._currentListName || "System").replace(/[^a-zA-Z0-9]/g, '').substring(0, 5);
        const pagePart = pageName.replace(/[^a-zA-Z0-9]/g, '').substring(0, 5);
        const guid = this.generateGuid();
        const errorId = `${listPart}_${pagePart}_${guid}`.toUpperCase();

        if (this._showAlerts && (severity === 'High' || severity === 'Medium')) {
            Swal.fire({
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
                confirmButtonText: 'Copy Error ID',
                confirmButtonColor: '#0078d4',
                preConfirm: () => {
                    const nav = navigator as any;
                    if (nav.clipboard && window.isSecureContext) {
                        return nav.clipboard.writeText(errorId)
                            .catch((err: any) => {
                                console.warn('[LoggerService] Clipboard API failed', err);
                                this.fallbackCopy(errorId);
                            });
                    } else {
                        this.fallbackCopy(errorId);
                    }
                }
            });
        }
        
        try {
            // Using the new SPFI instance to add the item
            await this._sp.web.lists.getByTitle(this._logListTitle).items.add({
                Title: this._currentListName,
                Page: fullUrl,
                ItemId: itemId,
                Module: module,
                Severity: severity,
                Error: errorMsg,
                ErrorId: errorId
            });
        } catch (e) {
            console.error('[LoggerService] CRITICAL: Failed to write to SharePoint list.', e);
        }
    }

    /**
     * GET LOGS FOR VIEWER
     */
    public static async getLogs(context: WebPartContext, sourceList: string): Promise<ILogItem[]> {
        if (!this._logListTitle || !this._sp) return [];
        try {
            // Using () to execute the query instead of .get()
            return await this._sp.web.lists.getByTitle(this._logListTitle).items
                .filter(`Title eq '${sourceList}'`)
                .select('Id', 'Title', 'Page', 'ItemId', 'Module', 'Severity', 'Error', 'ErrorId', 'Created', 'Author/Title')
                .expand('Author')
                .orderBy('Created', false)
                .top(5000)();
        } catch (e) {
            console.error("Could not fetch logs from list: " + this._logListTitle, e);
            return [];
        }
    }
}