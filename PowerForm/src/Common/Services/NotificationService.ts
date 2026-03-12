import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface INotificationRule {
  enabled: boolean;
  condition: string;
  message: string;
  targetGroups: any;
}

export interface INotificationConfig {
  enabled: boolean;

  // Standard Settings
  enableAdd: boolean; msgAdd: string; groupsAdd: number[];
  enableUpdate: boolean; msgUpdate: string; groupsUpdate: number[];
  enableDelete: boolean; msgDelete: string; groupsDelete: number[];
  enableView: boolean; msgView: string; groupsView: any[];

  // Rules Per Category
  rulesAdd: INotificationRule[];
  rulesUpdate: INotificationRule[];
  rulesDelete: INotificationRule[];
  rulesView: INotificationRule[];
}

export class NotificationService {
  private static listName = "PowerNotifications";
  private static _sp: SPFI;

  public static init(context: any): void {
    this._sp = spfi().using(SPFx(context));
  }

  public static async logNotification(
    siteUrl: string, 
    listName: string,
    action: 'Added' | 'Updated' | 'Deleted' | 'Viewed',
    itemData: any,
    currentUser: any,
    config: INotificationConfig
  ): Promise<void> {

    if (!config.enabled || !this._sp) return;

    let standardEnabled = false;
    let standardMsg = "";
    let standardGroups: any[] = [];
    let rules: INotificationRule[] = [];

    const userDisplay = currentUser ? (currentUser.Title || currentUser.LoginName) : "System";
    const userId = currentUser?.Id;

    switch (action) {
      case 'Added':
        standardEnabled = config.enableAdd;
        standardMsg = config.msgAdd || `{ListName} {Title} has been added by {userDisplay}.`;
        standardGroups = config.groupsAdd;
        rules = config.rulesAdd || [];
        break;
      case 'Updated':
        standardEnabled = config.enableUpdate;
        standardMsg = config.msgUpdate || `{ListName} {Title} has been updated by {userDisplay}.`;
        standardGroups = config.groupsUpdate;
        rules = config.rulesUpdate || [];
        break;
      case 'Deleted':
        standardEnabled = config.enableDelete;
        standardMsg = config.msgDelete || `{ListName} {Title} has been deleted by {userDisplay}.`;
        standardGroups = config.groupsDelete;
        rules = config.rulesDelete || [];
        break;
      case 'Viewed':
        standardEnabled = config.enableView;
        standardMsg = config.msgView || `{ListName} {Title} has been viewed by {userDisplay}.`;
        standardGroups = config.groupsView;
        rules = config.rulesView || [];
        break;
    }

    let ruleMatched = false;

    for (const rule of rules) {
      if (rule.enabled && rule.condition && rule.message) {
        const isMatch = this.evaluateCondition(rule.condition, itemData);

        if (isMatch) {
          ruleMatched = true;
          const finalMsg = this.formatMessage(rule.message, itemData, listName, userDisplay);

          let targetIds: number[] = [];
          if (rule.targetGroups) {
            if (Array.isArray(rule.targetGroups)) {
              targetIds = rule.targetGroups.map((g: any) => Number(g));
            } else if (typeof rule.targetGroups === 'string') {
              targetIds = rule.targetGroups.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n));
            }
          }

          await this.writeLog(siteUrl, listName, action, finalMsg, itemData, userId, targetIds);
        }
      }
    }

    if (!ruleMatched && standardEnabled) {
      const finalMsg = this.formatMessage(standardMsg, itemData, listName, userDisplay);
      const stdIds = (standardGroups || []).map((g: any) => {
        if (typeof g === 'number') return g;
        return parseInt(g.key || g.id);
      }).filter((n: any) => !isNaN(n));

      await this.writeLog(siteUrl, listName, action, finalMsg, itemData, userId, stdIds);
    }
  }

  private static evaluateCondition(condition: string, item: any): boolean {
    try {
      const evalItem = { ...item };
      Object.keys(item).forEach(key => {
        if (key.indexOf('OData_') === 0) {
          const cleanKey = key.replace('OData_', '');
          evalItem[cleanKey] = item[key];
        }
      });

      const jsExpr = condition
        .replace(/\seq\s/gi, " == ")
        .replace(/\sne\s/gi, " != ")
        .replace(/\sgt\s/gi, " > ")
        .replace(/\slt\s/gi, " < ")
        .replace(/\sge\s/gi, " >= ")
        .replace(/\sle\s/gi, " <= ")
        .replace(/\sand\s/gi, " && ")
        .replace(/\sor\s/gi, " || ");

      // eslint-disable-next-line no-new-func
      const func = new Function("item", `return ${jsExpr}`);
      return func(evalItem);
    } catch (e:any) {
      console.warn("[NotificationService] Syntax Error in Condition: " + condition);
      return false;
    }
  }

  private static formatMessage(template: string, item: any, listName: string, userDisplay: string): string {
    if (!template) return "";
    return template.replace(/\{([^}]+)\}/g, (match, key) => {
      const k = key.trim().toLowerCase();
      if (k === 'title') return item.Title || "Item";
      if (k === 'listname' || k === 'list name') return listName;
      if (k === 'id') return item.Id || "";
      if (['author', 'user', 'loggeduser', 'currentuser', 'userdisplay'].indexOf(k) > -1) return userDisplay;
      return item[key] || match;
    });
  }

  private static async writeLog(siteUrl: string, listName: string, action: string, msg: string, item: any, userId: number, targetGroups: number[]) {
    try {
      await this._sp.web.lists.getByTitle(this.listName).items.add({
        _Title: listName,
        _Message: msg,
        _Category: action,
        _Link: {
          Description: "View Item",
          Url: `${siteUrl}/lists/${listName}/DispForm.aspx?ID=${item.Id}`
        },
        _TargetGroupsId: { results: targetGroups || [] }
      });
    } catch (err:any) {
      console.error("[NotificationService] Failed to write to log list", err);
    }
  }
}