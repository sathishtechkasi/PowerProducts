import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import { ColumnType, IColumnDefinition } from './types';
import { WebPartContext } from "@microsoft/sp-webpart-base";

/**
 * Utility to generate CAML-based Field XML for Site Column creation.
 */
function generateFieldXml(col: IColumnDefinition): string {
  let xml = `<Field 
    ID="${col.id}" 
    Name="${col.internalName}" 
    DisplayName="${col.title}" 
    Type="${col.type}" 
    Group="${col.group}" 
    Required="${col.required ? 'TRUE' : 'FALSE'}"
    EnforceUniqueValues="${col.enforceUnique ? 'TRUE' : 'FALSE'}"
    StaticName="${col.internalName}"`;

  if (col.type === ColumnType.Text && col.maxLength) {
    xml += ` MaxLength="${col.maxLength}"`;
  }

  if (col.type === ColumnType.Choice || col.type === ColumnType.MultiChoice) {
    xml += ` Format="Dropdown">`;
    xml += `<CHOICES>`;
    (col.choices || []).forEach(choice => {
      xml += `<CHOICE>${choice}</CHOICE>`;
    });
    xml += `</CHOICES>`;
    if (col.type === ColumnType.MultiChoice) {
      xml += `<Default></Default>`;
    }
    xml += `</Field>`;
    return xml;
  }

  if (col.type === ColumnType.Lookup) {
    throw new Error("Use createLookupColumn for Lookup field creation");
  }

  xml += ` />`;
  return xml;
}

export class CommonService {
  private _sp: SPFI;
  private _siteUrl: string;

  constructor(siteUrl: string, context: WebPartContext) {
    this._siteUrl = siteUrl;
    // Initialize PnPjs v4 with the SPFx context
    this._sp = spfi(siteUrl).using(SPFx(context));
  }

  public get siteUrl(): string {
    return this._siteUrl;
  }

  /**
   * Cleans an item object of OData metadata and complex objects before sending to SP.
   */
  private _cleanItem(item: any): any {
    const cleaned: any = {};
    for (const key in item) {
      if (Object.prototype.hasOwnProperty.call(item, key)) {
        if (key.startsWith('__') || key.startsWith('OData_')) continue;
        
        const value = item[key];
        // Handle URL or Multi-value results objects
        if (value && typeof value === 'object') {
          if (value.Url && value.Description) {
            cleaned[key] = value;
          } else if (value.results) {
            cleaned[key] = value;
          }
          continue;
        }
        cleaned[key] = value;
      }
    }
    return cleaned;
  }

  public async getAllItems(listName: string): Promise<any[]> {
    // v4 uses the () execution syntax instead of .get()
    return this._sp.web.lists.getByTitle(listName).items.select("*")();
  }

  public async AddListItem(
    listName: string,
    item: any,
    attachmentsToAdd?: File[],
    attachmentsToDelete?: string[],
    updateItemId?: number
  ): Promise<number> {
    const list = this._sp.web.lists.getByTitle(listName);
    const cleanedItem = this._cleanItem(item);
    let itemId: number;

    if (updateItemId) {
      await list.items.getById(updateItemId).update(cleanedItem);
      itemId = updateItemId;
    } else {
      // In v4, items.add() returns the data object directly
      const result = await list.items.add(cleanedItem);
      itemId = result.Id;
    }

    const itemRef = list.items.getById(itemId);

    // Process deletions
    if (attachmentsToDelete?.length) {
      for (const fileName of attachmentsToDelete) {
        await itemRef.attachmentFiles.getByName(fileName).delete();
      }
    }

    // Process additions
    if (attachmentsToAdd?.length) {
      for (const file of attachmentsToAdd) {
        await itemRef.attachmentFiles.add(file.name, file);
      }
    }

    return itemId;
  }

  public async getLists(): Promise<{ key: string; text: string }[]> {
    const lists = await this._sp.web.lists
      .filter("Hidden eq false")
      .select("Title")();
    return lists.map(l => ({ key: l.Title, text: l.Title }));
  }

  public async getFieldsForList(listTitle: string): Promise<{ key: string; text: string }[]> {
    const fields = await this._sp.web.lists
      .getByTitle(listTitle)
      .fields.filter("Hidden eq false and ReadOnlyField eq false")
      .select("InternalName", "Title")();

    return fields
      .filter(f => f.InternalName !== 'ContentType')
      .map(f => ({ key: f.InternalName, text: f.Title }));
  }

  public async isValueUnique(listName: string, fieldName: string, value: string, excludeId?: number): Promise<boolean> {
    const safeVal = value.replace(/'/g, "''");
    let filter = `${fieldName} eq '${safeVal}'`;
    if (excludeId) filter += ` and Id ne ${excludeId}`;

    const items = await this._sp.web.lists
      .getByTitle(listName)
      .items.filter(filter)
      .select('Id')
      .top(1)();

    return items.length === 0;
  }

  public async getItemById(listName: string, itemId: number): Promise<any> {
    return this._sp.web.lists
      .getByTitle(listName)
      .items.getById(itemId)
      .select('*', 'Author/Title', 'Editor/Title')
      .expand('Author', 'Editor')();
  }

  public async createLookupColumn(col: IColumnDefinition): Promise<void> {
    if (!col.lookupListTitle || !col.lookupField) {
      throw new Error("lookupListTitle and lookupField are required");
    }

    const targetList = await this._sp.web.lists.getByTitle(col.lookupListTitle).select("Id")();
    const xml = `
      <Field 
        ID="${col.id}"
        Name="${col.internalName}" 
        StaticName="${col.internalName}" 
        DisplayName="${col.title}" 
        Type="Lookup" 
        List="${targetList.Id}" 
        ShowField="${col.showField || col.lookupField}" 
        Required="${col.required ? "TRUE" : "FALSE"}" 
        EnforceUniqueValues="${col.enforceUnique ? "TRUE" : "FALSE"}" 
        Mult="${col.isMulti ? "TRUE" : "FALSE"}" 
        Group="${col.group}" />`;

    await this._sp.web.fields.createFieldAsXml(xml);
  }

  public async createSiteColumn(col: IColumnDefinition): Promise<any> {
    if (col.type === ColumnType.Lookup) {
      return this.createLookupColumn(col);
    }
    const xml = generateFieldXml(col);
    return this._sp.web.fields.createFieldAsXml(xml);
  }

  public async stripHtml(html: string): Promise<string> {
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = html || '';
    return tempDiv.textContent || tempDiv.innerText || '';
  }
}