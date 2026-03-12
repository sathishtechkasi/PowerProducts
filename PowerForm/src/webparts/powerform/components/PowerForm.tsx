import * as React from 'react';
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import "@pnp/sp/site-groups";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/profiles";
import { Web } from "@pnp/sp/webs";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import Swal from 'sweetalert2';
import * as XLSX from 'xlsx';
import { IPowerFormProps, IRepeaterColumn } from './IPowerFormProps';
import { IPowerFormState, ColumnDefinition, ListItem, ILookupOption } from './IPowerFormState';
import styles from './PowerForm.module.scss';
import { CommonService } from '../../../Common/Services/CommonService';
import { NormalPeoplePicker, TagPicker, ITag } from '@fluentui/react/lib/Pickers';
import { ComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { ICustomAction } from './ICustomAction';
import { IViewConfig } from './ViewEditor';
import { LoggerService } from './LoggerService';
import { IChildListConfig } from './IPowerFormProps';
import { RepeaterInput } from './RepeaterInput';
import { NotificationService, INotificationConfig, INotificationRule } from '../../../Common/Services/NotificationService';

/**
 * Interface extensions for Office UI Fabric components 
 * to support additional metadata and data keys.
 */
interface IExtendedPersonaProps extends IPersonaProps {
  key: string;
  text?: string;
  secondaryText?: string;
}
interface IExtendedComboBoxOption extends IComboBoxOption {
  data?: any;
}
/**
 * SVG Icon library used for UI actions (Save, Edit, Delete, etc.)
 */
const Icons = {
  Plus: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M12 5v14M5 12h14" /></svg>,
  Export: () => <svg className={styles.icon} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z" /><polyline points="13 2 13 9 20 9" /><path d="M11 16h11" /><path d="m19 13 3 3-3 3" /></svg>,
  Search: () => <svg className={styles.icon} viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"></circle><path d="M21 21l-4.3-4.3"></path></svg>,
  Clear: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M18 6L6 18M6 6l12 12" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" /></svg>,
  Back: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M15 18l-6-6 6-6" /></svg>,
  Reset: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M4 4v6h6" /><path d="M20 20v-6h-6" /><path d="M5 19A9 9 0 1 1 19 5" /></svg>,
  Save: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M5 12l5 5L20 7" /></svg>,
  View: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z" /><circle cx="12" cy="12" r="3" /></svg>,
  Edit: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M12 20h9" /><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z" /></svg>,
  Delete: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M3 6h18" /><path d="M8 6V4h8v2" /><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6" /><path d="M10 11v6M14 11v6" /></svg>,
  SortAsc: () => <span style={{ fontSize: 10 }}> ▲</span>,
  SortDesc: () => <span style={{ fontSize: 10 }}> ▼</span>,
  Clip: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M16.5 6v11.5c0 2.21-1.79 4-4 4s-4-1.79-4-4V5a2.5 2.5 0 0 1 5 0v10.5c0 .55-.45 1-1 1s-1-.45-1-1V6H10v9.5a2.5 2.5 0 0 0 5 0V5c0-2.21-1.79-4-4-4S7 2.79 7 5v12.5c0 3.04 2.46 5.5 5.5 5.5s5.5-2.46 5.5-5.5V6h-1.5z" fill="currentColor" /></svg>,
  Refresh: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M17.65 6.35A7.958 7.958 0 0012 4c-4.42 0-7.99 3.58-7.99 8s3.57 8 7.99 8c3.73 0 6.84-2.55 7.73-6h-2.08A5.99 5.99 0 0112 18c-3.31 0-6-2.69-6-6s2.69-6 6-6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z" /></svg>,
  History: () => <svg className={styles.icon} viewBox="0 0 24 24"><path d="M12 2C6.5 2 2 6.5 2 12s4.5 10 10 10 10-4.5 10-10S17.5 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm.5-13H11v6l5.2 3.2.8-1.3-4.5-2.7V7z" /></svg>
};
/**
 * Custom Rich Text Editor component using document.execCommand.
 * Note: execCommand is deprecated but remains in use for IE11 compatibility.
 */
class RichTextEditor extends React.Component<any, any> {
  private editor!: HTMLDivElement;
  constructor(props: any) {
    super(props);
    this.state = { html: props.value || '' };
  }
  public componentDidUpdate(prevProps: any) {
    if (this.props.value !== prevProps.value && this.props.value !== this.state.html && document.activeElement !== this.editor) {
      this.setState({ html: this.props.value });
    }
  }
  // Toolbar Action Helper
  private exec = (cmd: string, val?: string) => {
    document.execCommand(cmd, false, val);
    this.handleChange();
  }
  private handleChange = () => {
    if (this.props.readOnly) return;
    const html = this.editor.innerHTML;
    this.setState({ html });
    if (this.props.onChange) this.props.onChange(html);
  }
  // Inside RichTextEditor render method
  public render() {
    const { readOnly } = this.props;
    const btnStyle: React.CSSProperties = {
      padding: '4px 8px',
      cursor: 'pointer',
      border: '1px solid #ccc',
      background: '#fff',
      borderRadius: '3px',
      fontWeight: 'bold' // Use string instead of number
    };
    return (
      <div style={{
        border: '1px solid #ccc',
        borderRadius: 4,
        background: readOnly ? '#f3f4f6' : '#fff',
        display: 'flex',           // Use flex to keep toolbar and area aligned
        flexDirection: 'column'
      }}>
        {!readOnly && (
          <div style={{ padding: '8px', borderBottom: '1px solid #eee', background: '#f9fafb', display: 'flex', gap: 4 }}>
            <button type="button" onClick={() => this.exec('bold')} style={{ fontWeight: 'bold', width: 24 }}>B</button>
            <button type="button" onClick={() => this.exec('italic')} style={{ fontStyle: 'italic', width: 24 }}>I</button>
            <button type="button" onClick={() => this.exec('underline')} style={{ textDecoration: 'underline', width: 24 }}>U</button>
            <span style={{ borderLeft: '1px solid #ccc', margin: '0 4px' }}></span>
            <button type="button" onClick={() => this.exec('justifyLeft')} style={{ width: 24 }}>L</button>
            <button type="button" onClick={() => this.exec('justifyCenter')} style={{ width: 24 }}>C</button>
            <button type="button" onClick={() => this.exec('justifyRight')} style={{ width: 24 }}>R</button>
            <span style={{ borderLeft: '1px solid #ccc', margin: '0 4px' }}></span>
            <button type="button" onClick={() => { const url = prompt('Enter Link URL:'); if (url) this.exec('createLink', url); }} style={{ width: 24 }}>🔗</button>
          </div>
        )}
        <div
          ref={el => { if (el) this.editor = el; }}
          contentEditable={!readOnly}
          style={{
            minHeight: 150,        // Increased starting height
            height: 'auto',        // ALLOW AUTO-EXPANSION 
            maxHeight: '600px',    // Prevent it from taking over the whole screen
            padding: 10,
            outline: 'none',
            overflowY: 'auto',     // Scrollbar only if it hits maxHeight
            background: readOnly ? '#f3f4f6' : '#fff',
            cursor: readOnly ? 'default' : 'text',
            resize: readOnly ? 'none' : 'vertical'
          }}
          dangerouslySetInnerHTML={{ __html: this.props.value }}
          onInput={this.handleChange}
          onBlur={this.handleChange}
        />
      </div>
    );
  }
}
/**
 * Main PowerForm component: Handles CRUD operations for SharePoint lists
 * dynamically based on list schema and configuration properties.
 */
export default class PowerForm extends React.Component<IPowerFormProps, IPowerFormState> {
  private _sp!: SPFI;
  private service: CommonService;
  constructor(props: IPowerFormProps) {
    super(props);
    // Parse Query Parameters (Modern Approach)
    const params = new URLSearchParams(window.location.search);
    const getParam = (key: string): string | null => {
      const lowerKey = key.toLowerCase();
      let result: string | null = null;
      params.forEach((value, key) => {
        if (key.toLowerCase() === lowerKey && result === null) result = value;
      });
      return result;
    };
    const rawQMode = getParam('qmode');
    const qMode = rawQMode ? rawQMode.toLowerCase() : 'list';
    const qId = getParam('itemid');
    let initialMode: 'list' | 'add' | 'edit' | 'view' = 'list';
    let initialId: number | undefined = undefined;
    // Check valid modes (case-insensitive)
    if (qMode) {
      const m = qMode.toLowerCase();
      if (m === 'add' || m === 'edit' || m === 'view' || m === 'list') {
        initialMode = m as any;
      }
    }
    // Check valid ID
    if (qId && !isNaN(Number(qId))) {
      initialId = Number(qId);
    }
    this.service = props.service;
    this.state = {
      fields: [],
      formData: {},
      loading: true,
      message: '',
      mode: initialMode,
      itemId: initialId,
      lookupOptions: {},
      formErrors: {},
      attachmentsNew: [],
      attachmentsDelete: [],
      existingAttachments: [],
      peopleOptions: {},
      attachments: [],
      canDelete: false,
      items: [],
      page: 1,
      pageSize: 10,
      totalItems: 0,
      selectedItems: [],
      canAdd: false,
      canEdit: false,
      canView: false,
      filters: {},
      searchText: '',
      sortField: 'Modified',
      sortDirection: 'desc',
      enableVersioning: true,
      autocompleteOptions: {},
      pickerSearch: {},
      activePickerKey: null,
      isPanelOpen: false,
      panelUrl: '',
      panelTitle: '',
      urlReadOnlyFields: [],
      currentUserGroups: [],
      pageCache: {},
      currentViewId: '',
      availableViews: [],
      activeViewFields: null,
      canSeeDefaultView: true,
      activeSectionIndex: 0,
      isBulkEditOpen: false,
      bulkEditField: '',
      bulkEditValue: null,
      isSaveDisabled: false,
      currentUser: null,
      childItems: {},
      isChildPanelOpen: false,
      activeChildConfig: null,
      activeChildItemIndex: -1,
      childFieldsCache: {}
    };
  }
  private indexedFields: string[] = [];
  private activeSearchStrategy: 'KQL' | 'INDEX' | 'CAML' = 'INDEX';
  private allSourceItems: any[] = [];
  //ADD: Remove listener on unmount
  public componentWillUnmount(): void {
    document.removeEventListener('mousedown', this.handleClickOutside);
  }
  // --- NOTIFICATION HELPERS ---


  //If columns start with underscore, we need to convert to OData format
  private toODataName(name: string): string {
    if (!name) return name;
    return name.indexOf('_') === 0 ? `OData_${name}` : name;
  }
  //  ADD: The Logic to close dropdowns
  private handleClickOutside = (event: any) => {
    if (this.state.activePickerKey) {
      // Find the wrapper of the currently open picker using the data attribute we will add
      const activeWrapper = document.querySelector(`[data-picker-wrapper="${this.state.activePickerKey}"]`);
      // If the click happened OUTSIDE the wrapper, close the picker
      if (activeWrapper && !activeWrapper.contains(event.target)) {
        this.setState({ activePickerKey: null });
      }
    }
  }
  //List View Configure
  //.aspx?LV={ViewName}&FilterField1=Number&FilterValue1={Number}
  //.aspx?LV=ViewName&FilterField1=Number&FilterValue1={Number}
  ///SitePages/Test4.aspx?LV=ViewName&FilterField1=Number&FilterValue1={Number}&FilterField2=Choice&FilterValue2={Choice}
  //SitePages/Test4.aspx?LV=ViewName&FilterField1=Number&FilterValue1={Number}&FilterField2=Choice&FilterValue2={Choice}
  public async componentDidMount(): Promise<void> {
    document.addEventListener('mousedown', this.handleClickOutside);
    //SETUP PNP
    this._sp = spfi().using(SPFx(this.props.context));
    try {
      const user = await this._sp.web.currentUser();

      this.setState({ currentUser: user });
      await this.loadFields();
      if (this.props.childConfigs && this.props.childConfigs.length > 0) {
        await this.loadChildSchemas();
      }
      //FETCH USER GROUPS 
      let currentGroupTitles: string[] = [];
      try {

        // Use getById(Id).groups.get() -> Safest method
        const groups = await this._sp.web.siteUsers.getById(this.state.currentUser.Id).groups();
        currentGroupTitles = groups.map((g: any) => g.Title);
        this.setState({ currentUserGroups: currentGroupTitles });
      } catch (error: any) {
        void LoggerService.log(
          'PowerForm - componentDidMount - Failed to fetch user groups (Proceeding without groups)',
          'High',
          this.state.itemId ? this.state.itemId.toString() : 'N/A',
          error.message || JSON.stringify(error)
        );
      }
      // CALCULATE AVAILABLE VIEWS & DEFAULT VIEW ACCESS
      const allViews = this.props.views || [];
      const allowedViews = allViews.filter((v: any) => {
        if (!v.allowedGroups || v.allowedGroups.length === 0) return true;
        return v.allowedGroups.some((g: string) => currentGroupTitles.indexOf(g) > -1);
      });
      // CHECK DEFAULT VIEW PERMISSIONS
      let canSeeDefault = true;
      const defGroups = this.props.defaultViewAllowedGroups || [];
      if (defGroups.length > 0) {
        // If groups are defined, user MUST be in one of them
        canSeeDefault = defGroups.some(g => currentGroupTitles.indexOf(g) > -1);
      }
      if (!canSeeDefault) {
      }
      // DETERMINE INITIAL VIEW
      const params = new URLSearchParams(window.location.search);
      const viewParam = params.get('LV');
      let initialViewId = ''; // Default to '' (All Items)
      let initialViewFields = null;
      let initialFilters = {};

      // LOGIC MATRIX:
      //  If URL param exists & allowed -> Use URL View
      //  Else If Can See Default -> Use Default ('')
      //  Else If Can See Custom Views -> Use First Allowed View
      //  Else -> ACCESS DENIED (No views available)
      if (viewParam) {
        const targetView = allowedViews.find((v: any) => v.title.toLowerCase() === viewParam.toLowerCase() || v.id === viewParam);
        if (targetView) {
          initialViewId = targetView.id;
          initialViewFields = targetView.visibleFields;
          // ... (Apply filters logic) ...
        }
      }
      else if (!canSeeDefault) {
        // User cannot see Default, so force them to the first available view
        if (allowedViews.length > 0) {
          initialViewId = allowedViews[0].id;
          initialViewFields = allowedViews[0].visibleFields;
          // Apply filters for this forced view
          const v = allowedViews[0];
          const initialFilters: { [key: string]: { value: any; operator: string } } = {};

          if (v.filters) {
            v.filters.forEach((f: any) => {
              // Check if field and value exist before assignment
              if (f.field && f.value) {
                // Accessing by f.field is now safe because of the index signature above
                initialFilters[f.field as string] = {
                  value: f.value,
                  operator: f.operator
                };
              }
            });
          }
        } else {
          // NO VIEWS AVAILABLE AT ALL
          this.setState({ loading: false, message: 'Access Denied: You do not have permission to view any lists.' });
          return; // STOP
        }
      }
      this.setState({
        availableViews: allowedViews,
        canSeeDefaultView: canSeeDefault, // Store for Render
        currentViewId: initialViewId,
        activeViewFields: initialViewFields,
        filters: initialFilters
      });
      //  CHECK LIST PERMISSIONS
      const perms = await this.checkPermissions();
      this.setState({
        canAdd: perms.canAdd,
        canEdit: perms.canEdit,
        canView: perms.canView,
        canDelete: perms.canDelete
      });
      //  VALIDATE URL & LOAD DATA
      const isAllowed = this.validateUrlPermissions(this.state.mode, perms);
      if (isAllowed) {
        if (this.state.mode === 'list') {
          await this.loadItems();
        }
        else if ((this.state.mode === 'edit' || this.state.mode === 'view') && this.state.itemId) {
          await this.loadItemData(this.state.itemId);
          if (this.state.mode === 'edit') {
            this.applyUrlParameters();
          }
        }
        else {
          this.applyUrlParameters();
          this.setState({ loading: false });
        }
        this.executeCustomScript(this.state.mode);
        // Trigger Cascades if needed
        if (this.props.cascadeConfig) {
          Object.keys(this.props.cascadeConfig).forEach(childKey => {
            const conf = this.props.cascadeConfig![childKey];
            const parentVal = this.state.formData[conf.parentField];
            if (parentVal) {
              void this.loadCascadeOptions(childKey, parentVal);
            }
          });
        }
      }
      if ((this.state.mode === 'view' || this.state.mode === 'edit') && this.state.itemId) {
        // Allow a small delay to ensure item data is fully loaded into state
        if (this.state.mode === 'view') {
          setTimeout(() => void this.triggerViewNotification(), 1000);
        }
      }
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - componentDidMount',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      this.setState({ message: 'Failed to initialize.', loading: false });
    }
  }
  private async triggerViewNotification() {
    // Optimization: Check if feature enabled before processing
    if (this.props.enableNotification) {
      try {
        const itemData = this.state.formData;
        if (!itemData || !itemData.Title) return;
        const config = this.getNotificationConfig();
        const listName = this.props.listPageTitle || this.props.selectedList;

        await NotificationService.logNotification(
          this.service.siteUrl,
          listName,
          'Viewed',
          { ...itemData, Id: this.state.itemId },
          this.state.currentUser,
          config
        );
      } catch (e: any) { console.error("View Notification Error", e); }
    }
  }


  // 2. Map Props to Config
  private getNotificationConfig(): INotificationConfig {
    return {
      enabled: this.props.enableNotification,

      // Add
      enableAdd: this.props.enableNotifAdd,
      msgAdd: this.props.msgNotifAdd,
      groupsAdd: this.props.groupsNotifAdd,
      rulesAdd: this.props.rulesAdd || [], // Pass Dynamic Array

      // Update
      enableUpdate: this.props.enableNotifUpdate,
      msgUpdate: this.props.msgNotifUpdate,
      groupsUpdate: this.props.groupsNotifUpdate,
      rulesUpdate: this.props.rulesUpdate || [], // Pass Dynamic Array

      // Delete
      enableDelete: this.props.enableNotifDelete,
      msgDelete: this.props.msgNotifDelete,
      groupsDelete: this.props.groupsNotifDelete,
      rulesDelete: this.props.rulesDelete || [], // Pass Dynamic Array

      // View
      enableView: this.props.enableNotifView,
      msgView: this.props.msgNotifView,
      groupsView: this.props.groupsNotifView,
      rulesView: this.props.rulesView || [] // Pass Dynamic Array
    };
  }

  // Add this helper class to your component

  private handleCustomAction(action: ICustomAction) {
    try {
      let url = action.url;
      const { itemId, listId } = this.state;
      // Regex to find all text inside curly braces: {Title}, {Status}, {MyLookup}, etc.
      url = url.replace(/\{([^}]+)\}/g, (match, placeholder) => {
        const key = placeholder.trim();
        const lowerKey = key.toLowerCase();
        //  Handle System Placeholders
        if (lowerKey === 'itemid') return itemId ? itemId.toString() : '';
        if (lowerKey === 'listid') return listId || '';
        if (lowerKey === 'siteurl') return this.service.siteUrl;
        //  Handle Form Fields (Dynamic)
        // Find the field definition by InternalName or EntityPropertyName
        let field = null;
        for (let i = 0; i < this.state.fields.length; i++) {
          const f = this.state.fields[i];
          if (f.InternalName === key || f.EntityPropertyName === key) {
            field = f;
            break; // Stop loop once found
          }
        }
        if (field) {
          // Get the LIVE value from the form state (works even if not saved yet)
          const val = this.state.formData[field.EntityPropertyName];
          // Handle Empty
          if (val === null || val === undefined) return '';
          // Handle Arrays (Multi-Choice, Multi-Lookup IDs) -> "1,2,3" or "A,B"
          if (Array.isArray(val)) {
            return encodeURIComponent(val.join(','));
          }
          // Handle Dates (Formatting ISO string)
          if ((field.TypeAsString === 'DateTime' || field.TypeAsString === 'Date') && val) {
            // You might want to grab just the YYYY-MM-DD part if strictly needed, 
            // but usually passing the full ISO string is safer.
            return encodeURIComponent(String(val));
          }
          // Handle Simple Values (Text, Number, Boolean, Single Lookup ID)
          return encodeURIComponent(String(val));
        }
        // If no match found, leave the placeholder as is (e.g. {UnknownParam})
        return match;
      });
      this.setState({
        isPanelOpen: true,
        panelUrl: url,
        panelTitle: action.title
      });
    }
    catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleCustomAction',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      void Swal.fire('Error', 'Failed to launch custom action.', 'error');
    }
  }
  private async checkPermissions(): Promise<{ canAdd: boolean; canEdit: boolean; canDelete: boolean; canView: boolean }> {
    try {
      //  Fetch List Metadata including the Permission Mask for the CURRENT USER
      // We use .select() to get the property safely without hitting the strict endpoint that caused 406 errors.
      const listData = await this._sp.web.lists
        .getByTitle(this.props.selectedList)
        .select('EffectiveBasePermissions')
        ();
      //  Extract High and Low values
      // PnP v2/v3 often returns them directly or inside a nested object depending on setup
      const perms = listData.EffectiveBasePermissions;
      if (!perms) {
        void LoggerService.log(
          'PowerForm - checkPermissions',
          'High',
          this.state.itemId ? this.state.itemId.toString() : 'N/A',
          'EffectiveBasePermissions not found in response. Defaulting to Visible.'
        );
        return { canAdd: true, canEdit: true, canDelete: true, canView: true };
      }
      //  The "StackExchange" Bitwise Logic
      // We use the 'Low' value because Add/Edit/Delete/View bits are all in the lower range.
      // ViewListItems   = 1  (0x00000001)
      // AddListItems    = 2  (0x00000002)
      // EditListItems   = 4  (0x00000004)
      // DeleteListItems = 8  (0x00000008)
      const low = perms.Low;
      // Helper function to check if a specific bit is set
      // (val & mask) === mask
      const hasRight = (mask: number) => (low & mask) === mask;
      const canView = hasRight(1);
      const canAdd = hasRight(2);
      const canEdit = hasRight(4);
      const canDelete = hasRight(8);
      return { canAdd, canEdit, canDelete, canView };
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - checkPermissions',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      // Fallback: Default to TRUE so the UI doesn't break
      return { canAdd: false, canEdit: false, canDelete: false, canView: true };
    }
  }
  public componentDidUpdate(prevProps: IPowerFormProps, prevState: IPowerFormState) {
    try {

      //  Handle Mode Switching  
      if (prevState.mode !== this.state.mode) {
        this.executeCustomScript(this.state.mode);
        //  If user clicks "Back" to go to List, but list is empty (because we deep-linked), load it now.
        if (this.state.mode === 'list' && this.state.items.length === 0) {
          void this.loadItems();
        }
      }
      if (prevProps.isLargeList !== this.props.isLargeList) {
        console.log(`[PowerForm] Large List Mode changed to: ${this.props.isLargeList}. Reloading...`);

        // Clear internal caches to prevent stale data
        this.allSourceItems = [];

        this.setState({
          items: [],
          page: 1,
          searchText: '',
          filters: {},
          nextPageUrl: undefined,
          pageCache: {},
          loading: true
        }, () => {
          // Reload using the NEW mode (The loadItems method checks this.props.isLargeList)
          void this.loadItems('init');
        });
      }
      if (JSON.stringify(prevProps.childConfigs) !== JSON.stringify(this.props.childConfigs)) {
        void this.loadChildSchemas();
      }
      //  NEW: Detect Lookup Config Changes (Property Pane updates)
      // If you add a Filter Query or change columns in the panel, this forces the lookups to reload immediately.
      if (JSON.stringify(prevProps.lookupDisplayConfig) !== JSON.stringify(this.props.lookupDisplayConfig)) {
        this.state.fields.forEach(f => {
          // Re-fetch options for any Lookup field to apply the new Filter Query
          if (f.TypeAsString.indexOf('Lookup') > -1) {
            void this.fetchLookupOptions(f);
          }
        });
      }
      // NEW: Detect Cascade Config Changes
      if (JSON.stringify(prevProps.cascadeConfig) !== JSON.stringify(this.props.cascadeConfig)) {
        // We don't force-fetch here because cascades wait for a parent selection,
        // but this ensures the next selection triggers the correct logic.
      }
    }
    catch (error: any) {
      void LoggerService.log(
        'PowerForm - componentDidUpdate',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private onAutocompleteSearch = async (text: string, fieldKey: string): Promise<IExtendedComboBoxOption[]> => {
    //  GET CONFIG
    const config = this.props.autocompleteConfig ? this.props.autocompleteConfig[fieldKey] : null;
    if (!config || !config.sourceList || !config.sourceField) return [];
    try {
      const isGuid = config.sourceList.match(/^[0-9a-f]{8}-/i);
      const list = isGuid
        ? this._sp.web.lists.getById(config.sourceList)
        : this._sp.web.lists.getByTitle(config.sourceList);
      // Convert Source Field to OData Name if needed
      const sourceFieldOData = this.toODataName(config.sourceField);
      let filterQuery = '';
      //  BUILD SEARCH FILTER (Use OData Name)
      if (text) {
        filterQuery = `substringof('${encodeURIComponent(text)}', ${sourceFieldOData})`;
      }
      if (config.sourceQuery) {
        filterQuery = filterQuery
          ? `(${filterQuery}) and (${config.sourceQuery})`
          : config.sourceQuery;
      }
      //  SMART SELECT (Use OData Names)
      const selectFields = ['Id', sourceFieldOData];
      // A. Add Additional Display Fields
      if (config.additionalFields) {
        config.additionalFields.forEach(f => {
          const fOData = this.toODataName(f);
          if (selectFields.indexOf(fOData) === -1) selectFields.push(fOData);
        });
      }
      // B. Add Mapped Source Fields
      if (config.columnMapping) {
        config.columnMapping.forEach(map => {
          if (map.source) {
            const fOData = this.toODataName(map.source);
            if (selectFields.indexOf(fOData) === -1) selectFields.push(fOData);
          }
        });
      }
      //  EXECUTE QUERY
      let query = list.items.select(...selectFields).top(20);
      if (filterQuery) {
        query = query.filter(filterQuery);
      }
      const items = await query();
      //  MAP RESULTS
      const uniqueVals = new Set<string>();
      const options: IExtendedComboBoxOption[] = [];
      items.forEach((i: any) => {
        //  Try reading with InternalName OR OData Name
        const val = i[config.sourceField] || i[sourceFieldOData];
        if (val && !uniqueVals.has(val)) {
          uniqueVals.add(val);
          options.push({ key: val, text: val, data: i } as any);
        }
      });
      return options;
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - onAutocompleteSearch',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return [];
    }
  }
  private openVersionHistory = () => {
    try {
      // Get State
      const { listId, listUrl, itemId } = this.state;
      // Error Handling: Alert user instead of failing silently
      if (!itemId) {
        // ADDED: Log specific warning if ID is missing
        void LoggerService.log(
          'PowerForm - openVersionHistory',
          'Medium',
          'N/A',
          'Attempted to open version history without a valid Item ID.'
        );
        return;
      }
      // 3. Fallback for List ID (If loadFields failed to capture it, try using the context)
      const safeListId = listId || this.props.selectedList;
      // 4. Construct URL
      // We wrap safeListId in curly braces only if it looks like a GUID
      const listParam = /^[0-9a-fA-F-]{36}$/.test(String(safeListId)) ? `{${safeListId}}` : safeListId;
      const fileName = `${listUrl}/${itemId}_.000`;
      const encodedFileName = encodeURIComponent(fileName);
      // Construct URL
      const historyUrl = `${this.service.siteUrl}/_layouts/15/Versions.aspx?list=${listParam}&ID=${itemId}&FileName=${encodedFileName}&IsDlg=1`;
      // 5. Open Panel
      this.setState({
        isPanelOpen: true,
        panelUrl: historyUrl,
        panelTitle: 'Version History'
      });
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - openVersionHistory',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private async loadCascadeOptions(childFieldName: string, parentValue: any): Promise<void> {
    if (!parentValue) {
      this.setState(prev => ({ lookupOptions: { ...prev.lookupOptions, [childFieldName]: [] } }));
      return;
    }
    let parentId = parentValue;
    if (typeof parentValue === 'object' && parentValue.Id) parentId = parentValue.Id;
    const rawConfig = this.props.cascadeConfig ? this.props.cascadeConfig[childFieldName] : null;
    const config: any = rawConfig;
    if (!config) return;
    const fieldDef = this.state.fields.find(f => f.InternalName === childFieldName);
    if (!fieldDef || !fieldDef.LookupList) return;
    let listId = fieldDef.LookupList.replace(/^{|}$/g, '');
    try {
      const cascadeFilter = `${config.foreignKey}/Id eq ${parentId}`;
      let finalFilter = cascadeFilter;
      if (config.filterQuery) {
        finalFilter = `(${cascadeFilter}) and (${config.filterQuery})`;
      }
      const selectFields = ['Id', 'Title'];
      //  Convert Additional Fields
      if (config.additionalFields) {
        // A. Additional Fields
        if (config.additionalFields) {
          config.additionalFields.forEach((f: string) => {
            const fOData = this.toODataName(f);
            if (selectFields.indexOf(fOData) === -1) selectFields.push(fOData);
          });
        }
        // B. Mapped Fields 
        if (config.columnMapping) {
          config.columnMapping.forEach((map: any) => {
            if (map.source) {
              const fOData = this.toODataName(map.source);
              if (selectFields.indexOf(fOData) === -1) selectFields.push(fOData);
            }
          });
        }
      }
      // PnPjs: Fetch Cascade Options
      const items = await this._sp.web.lists.getById(listId).items
        .select(...selectFields)
        .filter(finalFilter)
        .top(5000)();
        
      const options: ILookupOption[] = items.map((item: any) => ({
        key: item.Id,
        text: item.Title,
        itemData: item
      }));
      this.setState(prev => ({
        lookupOptions: { ...prev.lookupOptions, [childFieldName]: options }
      }));
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - loadCascadeOptions',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private async loadFields(): Promise<void> {
    try {
      // PnPjs: Fetch list metadata
      const listData = await this._sp.web.lists.getByTitle(this.props.selectedList)
        .select('Id', 'EnableVersioning', 'RootFolder/ServerRelativeUrl', 'RootFolder/Name')
        .expand('RootFolder')();

      const manualListUrl = `${this.service.siteUrl}/Lists/${this.props.selectedList}`;

      this.setState({
        listId: listData.Id,
        enableVersioning: listData.EnableVersioning,
        listUrl: manualListUrl
      });

      // PnPjs: Fetch list fields
      const fieldsData = await this._sp.web.lists.getByTitle(this.props.selectedList).fields
        .filter("Hidden eq false and (ReadOnlyField eq false or InternalName eq 'Attachments')")();

      this.setState({ fields: fieldsData });
      // Wait for all lookups to load so URL parameters can match text values
      const lookupPromises: Promise<void>[] = [];
     for (const field of fieldsData) {
        // Cast to 'any' because LookupList is not in the base IFieldInfo interface
        if (field.FieldTypeKind === 7 && (field as any).LookupList) {

          lookupPromises.push(this.fetchLookupOptions(field));
        }
      }
      if (lookupPromises.length > 0) {
        await Promise.all(lookupPromises);
      }
    } catch (error: any) {
      this.setState({ message: 'Failed to load fields.', loading: false });
      void LoggerService.log(
        'PowerForm - loadFields',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private async loadChildSchemas(): Promise<void> {
    const configs = this.props.childConfigs || [];
    const cache: any = {};

    for (const conf of configs) {
      if (!conf.childListTitle) continue;
      try {
        // PnPjs: Fetch fields for the child list
        const fieldsData = await this._sp.web.lists.getByTitle(conf.childListTitle).fields
          .filter("Hidden eq false and ReadOnlyField eq false")();
          
        cache[conf.childListTitle] = fieldsData;
      } catch (e: any) {
        console.error("Error loading child list schema", e);
      }
    }
    this.setState({ childFieldsCache: cache });
  }
  /**
   * Parses URL query parameters and pre-fills the form.
   * Supports: ?Title=Test, ?Amount=100, ?CountryId=5, ?Country=Canada
   */
  /**
   * Parses URL query parameters and pre-fills the form.
   * Uses standard loops to be IE11 compatible (No .find).
   */
  private applyUrlParameters(): void {
    try {
      const params = new URLSearchParams(window.location.search);
      const newFormData = { ...this.state.formData };
      const lockedFields: string[] = [];
      let hasChanges = false;

      params.forEach((rawVal, rawKey) => {
        const paramKey = decodeURIComponent(rawKey);
        const paramVal = decodeURIComponent(rawVal);

        // Skip system params (return acts as 'continue' inside a forEach)
        if (paramKey.toLowerCase() === 'qmode' || paramKey.toLowerCase() === 'itemid') return;

        let field = this.state.fields.find(f => f.EntityPropertyName === paramKey || f.InternalName === paramKey);
        let isIdReference = false;

        // Check if param is an ID reference (e.g. CountryId -> Country)
        if (!field && paramKey.toLowerCase().endsWith('id')) {
          const baseName = paramKey.substring(0, paramKey.length - 2);
          field = this.state.fields.find(f => f.EntityPropertyName === baseName || f.InternalName === baseName);
          if (field) isIdReference = true;
        }

        if (field) {
          const key = field.EntityPropertyName;
          let finalVal: any = paramVal;
          const type = field.TypeAsString;

          if (type === 'Number' || type === 'Currency') {
            finalVal = parseFloat(paramVal);
            if (isNaN(finalVal)) finalVal = null;
          } else if (type === 'Boolean') {
            finalVal = (paramVal.toLowerCase() === 'true' || paramVal === '1');
          } else if (type === 'Lookup' || type === 'User' || type === 'LookupMulti' || type === 'UserMulti') {
            const numVal = parseInt(paramVal, 10);
            if (!isNaN(numVal)) {
              finalVal = numVal;
            } else if (!isIdReference) {
              const opts = this.state.lookupOptions[key] || [];
              const match = opts.find(o => o.text && o.text.toLowerCase() === paramVal.toLowerCase());
              if (match) {
                finalVal = Number(match.key);
              } else {
                void LoggerService.log('PowerForm', 'Medium', 'N/A', `URL Param: No lookup match for '${paramVal}'`);
                finalVal = null;
              }
            }
            if ((type === 'LookupMulti' || type === 'UserMulti') && finalVal !== null) finalVal = [finalVal];
          } else if (type === 'MultiChoice') {
            finalVal = paramVal.split(',').map(s => s.trim());
          }

          if (finalVal !== null && finalVal !== undefined) {
            newFormData[key] = finalVal;
            lockedFields.push(key);
            hasChanges = true;
          }
        }
      });

      if (hasChanges) {
        this.setState({ formData: newFormData, urlReadOnlyFields: lockedFields }, () => {
          if (this.props.cascadeConfig) {
            Object.keys(this.props.cascadeConfig).forEach(childKey => {
              const conf = this.props.cascadeConfig![childKey];
              const parentVal = this.state.formData[conf.parentField];
              if (parentVal) void this.loadCascadeOptions(childKey, parentVal);
            });
          }
        });
      }
    } catch (error: any) {
      void LoggerService.log('PowerForm - applyUrlParameters', 'Medium', 'N/A', error.message);
    }
  }
  private async fetchLookupOptions(field: ColumnDefinition): Promise<void> {
    const fieldName = field.EntityPropertyName;
    let listId = field.LookupList || '';
    listId = listId.replace(/^{|}$/g, '');
    const config = this.props.lookupDisplayConfig ? this.props.lookupDisplayConfig[field.InternalName] || this.props.lookupDisplayConfig[field.EntityPropertyName] : null;
    try {
      let selectFields = ['Id', 'Title'];
      //  Convert Additional Fields to OData Name
      if (config && config.additionalFields && config.additionalFields.length > 0) {

        if (config.additionalFields) {
          config.additionalFields.forEach((f: string) => {
            const fOData = this.toODataName(f);
            if (selectFields.indexOf(fOData) === -1) selectFields.push(fOData);
          });
        }
        if (config.columnMapping) {
          config.columnMapping.forEach((map: any) => {
            if (map.source) {
              const fOData = this.toODataName(map.source);
              if (selectFields.indexOf(fOData) === -1) selectFields.push(fOData);
            }
          });
        }
      }
      // PnPjs: Fetch Lookup Options
      let query = this._sp.web.lists.getById(listId).items.select(...selectFields).top(5000);
      if (config && config.filterQuery) {
        query = query.filter(config.filterQuery);
      }
      
      const items = await query();
      
      const options: ILookupOption[] = items.map((item: any) => ({
        key: item.Id,
        text: item.Title,
        itemData: item
      }));
      this.setState(prev => ({
        lookupOptions: { ...prev.lookupOptions, [fieldName]: options }
      }));
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - fetchLookupOptions',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private async loadItems(mode: 'init' | 'search' | 'more' | 'filter' = 'init'): Promise<void> {
    // Check the boolean passed from Web Part Properties
    if (this.props.isLargeList) {
      return this.loadItems_above5K(mode);
    } else {
      return this.loadItems_upto5K();
    }
  }

  private async loadItems_upto5K(): Promise<void> {
    const limit = 5000;
    let currentFields = this.state.fields;
    // 1. Ensure Fields are Loaded
    if (!currentFields || currentFields.length === 0) {
      try {
        currentFields = await this._sp.web.lists.getByTitle(this.props.selectedList)
          .fields.filter("Hidden eq false and ReadOnlyField eq false")();
        this.setState({ fields: currentFields });
      } catch (error: any) {
        void LoggerService.log('PowerForm - loadItems_upto5K (Fields)', 'High', 'N/A', error.message);
      }
    }
    if (this.allSourceItems.length === 0) {
      this.setState({ loading: true });
      try {
        const list = this._sp.web.lists.getByTitle(this.props.selectedList);
        const listCols = this.props.listVisibleFields && this.props.listVisibleFields.length > 0
          ? [...this.props.listVisibleFields]
          : (currentFields ? currentFields.filter(f => f.InternalName !== 'ContentType').map(f => f.InternalName) : []);
        const selectSet = new Set<string>();
        const expandSet = new Set<string>();
        // BASE SYSTEM FIELDS
        ['Id', 'Attachments', 'Modified', 'Created'].forEach(c => selectSet.add(c));
        ['Author', 'Editor'].forEach(c => {
          expandSet.add(c);
          selectSet.add(c + '/Title');
          selectSet.add(c + '/EMail');
        });
        // 2. DYNAMIC FIELD LOGIC (Visible Columns Only)
        currentFields.forEach(field => {
          const entity = field.EntityPropertyName;
          if (listCols.indexOf(field.InternalName) > -1 || listCols.indexOf(entity) > -1) {
            const type = field.TypeAsString || '';
            const isLookup = type.indexOf('Lookup') > -1;
            const isUser = type.indexOf('User') > -1;
            if (isLookup || isUser) {
              expandSet.add(entity);
              selectSet.add(entity + '/Id');
              selectSet.add(entity + '/Title');
            } else {
              selectSet.add(entity);
            }
          }
        });
        // 3.Convert Sets to Arrays manually for ES5 compatibility
        const selectFields: string[] = [];
        selectSet.forEach(s => selectFields.push(s));
        const expandFields: string[] = [];
        expandSet.forEach(e => expandFields.push(e));
        // 4. FETCH
        const rawItems = await list.items
          .select(...selectFields)
          .expand(...expandFields)
          .orderBy('Modified', false)
          .top(limit)
          ();
        this.allSourceItems = rawItems.map(item => ({
          ...item,
          Id: item.Id || 0,
          Attachments: item.Attachments || false
        }));
        this.applyClientSideFilters();
        this.setState({ loading: false }, () => {
          if (this.state.mode === 'list') this.executeCustomScript('list');
        });
      } catch (error: any) {
        void LoggerService.log('PowerForm - loadItems_upto5K', 'High', 'N/A', error.message);
        this.setState({ loading: false });
      }
    } else {
      this.applyClientSideFilters();
    }
  }
  private async ensureIndexedColumns(): Promise<void> {
    if (this.indexedFields.length > 0) return; // Already loaded
    try {
      // PnPjs: Ensure indexed columns
      const fieldsData = await this._sp.web.lists.getByTitle(this.props.selectedList).fields
        .select('InternalName', 'Indexed')
        .filter('Indexed eq true')();
        
      this.indexedFields = fieldsData.map((f: any) => f.InternalName);
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - ensureIndexedColumns',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private async loadItems_above5K(mode: 'init' | 'search' | 'more' | 'filter' | 'page' = 'init'): Promise<void> {
    const listTitle = this.props.selectedList;
    const listId = this.state.listId;
    const pageSize = this.state.pageSize || 30;
    console.group(`[LargeList] loadItems_above5K | Mode: ${mode}`);
    //  Reset State (if not paging)
    if (mode === 'init' || mode === 'search' || mode === 'filter') {
      this.setState({ loading: true, items: [], page: 1, nextPageUrl: undefined, pageCache: {} });
      // Reset Strategy to default for new searches
      if (mode === 'search') this.activeSearchStrategy = 'KQL'; // Start with KQL
      else this.activeSearchStrategy = 'INDEX'; // Default for Init/Filter
    } else {
      this.setState({ loading: true });
    }
    // HELPER: Define the 3 Search Methods
    // --- STRATEGY A: KQL ---
    const runKQL = async (): Promise<boolean> => {
      try {
        console.log(" Attempt 1: KQL (Search API)...");
        let startRow = 0;
        if (mode === 'page' && this.state.nextPageUrl) {
          // Parse our fake KQL token
          const parts = this.state.nextPageUrl.split('StartRow=');
          if (parts[1]) startRow = parseInt(parts[1]);
        }
        const kqlQuery = `'${this.state.searchText}* AND ListId:${listId}'`;
        const selectProps = "'ListItemID,Title,Author,Editor,Created,Write,EditorOWSUSER,AuthorOWSUSER,Path,OriginalPath'";
        // --- ADD SORTING TO KQL (Write:descending) ---
        const searchUrl = `${this.service.siteUrl}/_api/search/query?querytext=${kqlQuery}&selectproperties=${selectProps}&rowlimit=${pageSize}&startrow=${startRow}&clienttype='ContentSearchRegular'&sortlist='Write:descending'`;
        const response = await this.props.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        if (!response.ok) return false; // Fail to next strategy
        const json = await response.json();
        let resultTable: any[] = [];
        let totalRows = 0;
        if (json.PrimaryQueryResult && json.PrimaryQueryResult.RelevantResults) {
          totalRows = json.PrimaryQueryResult.RelevantResults.TotalRows;
          if (json.PrimaryQueryResult.RelevantResults.Table && json.PrimaryQueryResult.RelevantResults.Table.Rows) {
            resultTable = json.PrimaryQueryResult.RelevantResults.Table.Rows;
          }
        }
        // WATERFALL CHECK: If 0 results, return FALSE to trigger fallback
        if (totalRows === 0 && startRow === 0) {
          void LoggerService.log(
            'PowerForm - runKQL -KQL Error. Falling back',
            'Medium',
            this.state.itemId ? this.state.itemId.toString() : 'N/A',
            'KQL returned 0 results. Falling back...'
          );
          return false;
        }
        // Map Results
        const mappedItems = resultTable.map((row: any) => {
          const item: any = {};
          row.Cells.forEach((cell: any) => {
            if (cell.Key === 'ListItemID') item.Id = parseInt(cell.Value, 10);
            else if (cell.Key === 'Title') item.Title = cell.Value;
            else if (cell.Key === 'Path') item.FileRef = cell.Value;
            else if (cell.Key === 'Author') item.Author = { Title: cell.Value };
            else if (cell.Key === 'Editor') item.Editor = { Title: cell.Value };
            else if (cell.Key === 'Write') item.Modified = cell.Value;
          });
          item.Attachments = row.Cells.some((c: any) => c.Key === 'Attachments' && (c.Value === '1' || c.Value === 'true'));
          return item;
        });
        let nextKqlUrl = null;
        if ((startRow + mappedItems.length) < totalRows) {
          nextKqlUrl = `KQL?StartRow=${startRow + pageSize}`;
        }
        this.activeSearchStrategy = 'KQL'; // Confirm this as the winner
        this.updateState(mappedItems, nextKqlUrl ?? "", totalRows, mode);
        return true; // Success
      } catch (error: any) {
        void LoggerService.log(
          'PowerForm - runKQL -KQL Error. Falling back',
          'Medium',
          this.state.itemId ? this.state.itemId.toString() : 'N/A',
          error.message || JSON.stringify(error)
        );
        return false;
      }
    };
    // --- STRATEGY B & C: RENDER LIST DATA (INDEX or CAML) ---
    const runRenderListData = async (useCaml: boolean): Promise<boolean> => {
      const stratName = useCaml ? "CAML (Live)" : "INDEX (Standard)";
      console.log(` Attempt ${useCaml ? '3' : '2'}: ${stratName}...`);
      // Show Warning if falling back to CAML
      if (useCaml) {
        void Swal.fire({
          toast: true,
          position: 'top-end',
          icon: 'warning',
          title: 'Deep Search Active',
          text: 'Fast search returned 0 results. Scanning all items (this may be slow)...',
          showConfirmButton: false,
          timer: 4000
        });
      }
      let viewFieldsXml = '<FieldRef Name="ID" /><FieldRef Name="Title" /><FieldRef Name="Attachments" /><FieldRef Name="Modified" /><FieldRef Name="Created" /><FieldRef Name="Author" /><FieldRef Name="Editor" /><FieldRef Name="FileRef" />';
      const listCols = this.props.listVisibleFields || [];
      const fieldsToCheck = this.state.fields || [];
      listCols.forEach(col => { if (viewFieldsXml.indexOf(`Name="${col}"`) === -1) viewFieldsXml += `<FieldRef Name="${col}" />`; });
      let endpoint = `${this.service.siteUrl}/_api/web/lists/getbytitle('${listTitle}')/RenderListDataAsStream`;
      let viewXmlQuery = '';
      // --- BUILD QUERY ---
      if (!useCaml) {
        // INDEXED MODE
        const safeSearch = encodeURIComponent(this.state.searchText);
        endpoint += `?InplaceSearchQuery=${safeSearch}`;
        const filterXml = this.buildFilterXml_Strict();
        if (filterXml) viewXmlQuery = `<Where>${filterXml}</Where>`;
      } else {
        // CAML MODE (Deep Search)
        let searchXml = '';
        if (this.state.searchText) {
          const val = this.state.searchText;
          // ... (Your Existing CAML Builder Logic) ...
          const searchClauses: string[] = [];
          const colsToSearch = [...listCols];
          if (colsToSearch.indexOf('Title') === -1) colsToSearch.push('Title');
          colsToSearch.forEach(key => {
            if (['Attachments', 'ContentType', 'FileRef', 'Created', 'Modified'].indexOf(key) > -1) return;
            const fieldDef = fieldsToCheck.find(f => f.InternalName === key);
            const type = fieldDef ? fieldDef.TypeAsString : 'Text';
            if (type === 'Number' || type === 'Currency') {
              const numVal = parseFloat(val);
              if (!isNaN(numVal)) searchClauses.push(`<Eq><FieldRef Name="${key}"/><Value Type="Number">${numVal}</Value></Eq>`);
            } else {
              searchClauses.push(`<Contains><FieldRef Name="${key}"/><Value Type="Text">${val}</Value></Contains>`);
            }
          });
          if (searchClauses.length > 0) {
            if (searchClauses.length === 1) searchXml = searchClauses[0];
            else {
              let combined = searchClauses[0];
              for (let i = 1; i < searchClauses.length; i++) combined = `<Or>${combined}${searchClauses[i]}</Or>`;
              searchXml = combined;
            }
          }
        }
        const filterXml = this.buildFilterXml_Loose();
        // Combine Search + Filters
        if (searchXml && filterXml) viewXmlQuery = `<Where><And>${searchXml}${filterXml}</And></Where>`;
        else if (searchXml) viewXmlQuery = `<Where>${searchXml}</Where>`;
        else if (filterXml) viewXmlQuery = `<Where>${filterXml}</Where>`;
      }
      // -- ADD SORTING TO CAML/INDEX (Last Modified Top) ---
      const orderByXml = `<OrderBy><FieldRef Name="Modified" Ascending="FALSE" /></OrderBy>`;
      // Paging Token
      let pagingParam = "";
      if ((mode === 'more' || mode === 'page') && this.state.nextPageUrl) {
        const parts = this.state.nextPageUrl.split('?');
        pagingParam = parts.length > 1 ? parts[1] : parts[0];
      }
      const payload = {
        parameters: {
          RenderOptions: 5707527,
          // Inject orderByXml BEFORE viewXmlQuery (which is the Where clause)
          ViewXml: `<View Scope="RecursiveAll"><ViewFields>${viewFieldsXml}</ViewFields><RowLimit Paged="TRUE">${pageSize}</RowLimit><Query>${orderByXml}${viewXmlQuery}</Query></View>`,
          DatesInUtc: true,
          Paging: pagingParam
        }
      };
      try {
        const response = await this.props.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { body: JSON.stringify(payload) });
        if (!response.ok) return false;
        const data = await response.json();
        const rawRows = (data.ListData && data.ListData.Row) ? data.ListData.Row : [];
        // WATERFALL CHECK: If 0 results, return FALSE (unless we are already in CAML mode)
        if (rawRows.length === 0 && !useCaml && mode !== 'page') {
          void LoggerService.log(
            'PowerForm - runRenderListData',
            'Medium',
            this.state.itemId ? this.state.itemId.toString() : 'N/A',
            "-> Index Search returned 0 results. Falling back to CAML..."
          );
          return false;
        }
        const mappedItems = rawRows.map((row: any) => ({ ...row, Id: parseInt(row.ID, 10), Attachments: row.Attachments === "1" || row.Attachments === "true" }));
        let finalTotal = 0;
        if (data.ListData && data.ListData.FilterLink) {
          const u = new URLSearchParams(data.ListData.FilterLink);
          if (u.get('ViewCount')) finalTotal = parseInt(u.get('ViewCount') || "0");
        }
        if (finalTotal === 0) {
          if (data.ItemCount) {
            finalTotal = parseInt(data.ItemCount);
          } else {
            try {
              // Fallback: Get real count from list metadata

              const listData = await this._sp.web.lists.getByTitle(listTitle).select('ItemCount')();
              finalTotal = listData.ItemCount;
            } catch (e: any) {
              finalTotal = 0;
            }
          }
        } if (useCaml && this.state.searchText) finalTotal = mappedItems.length; // Approximate for CAML
        this.activeSearchStrategy = useCaml ? 'CAML' : 'INDEX';
        this.updateState(mappedItems, data.ListData ? data.ListData.NextHref : null, finalTotal, mode);
        return true;
      } catch (error: any) {
        void LoggerService.log(
          'PowerForm - runRenderListData',
          'Medium',
          this.state.itemId ? this.state.itemId.toString() : 'N/A',
          error.message || JSON.stringify(error)
        );
        return false;
      }
    };
    // =========================================================
    // EXECUTION FLOW (The Waterfall)
    // =========================================================
    // CASE 1: PAGINATION (Stick to the winning strategy)
    if (mode === 'page' || mode === 'more') {
      if (this.activeSearchStrategy === 'KQL') {
        await runKQL();
      } else if (this.activeSearchStrategy === 'CAML') {
        await runRenderListData(true);
      } else {
        await runRenderListData(false);
      }
      console.groupEnd();
      return;
    }
    // CASE 2: SEARCH (Try 1 -> 2 -> 3)
    if (mode === 'search' && this.state.searchText) {
      // Try KQL
      const kqlSuccess = await runKQL();
      if (kqlSuccess) { console.groupEnd(); return; }
      // Try INDEX
      const indexSuccess = await runRenderListData(false);
      if (indexSuccess) { console.groupEnd(); return; }
      // Try CAML (Final Fallback)
      await runRenderListData(true);
      console.groupEnd();
      return;
    }
    // CASE 3: DEFAULT (Init / Filter - Use Standard Index)
    await runRenderListData(false);
    console.groupEnd();
  }
  // Helper to update state uniformly
  private updateState(items: any[], nextHref: string, totalItems: number, mode: string) {
    this.setState(prev => {
      const newCache = { ...prev.pageCache };
      newCache[prev.page] = { items: items, nextHref: nextHref };
      return {
        items: items, nextPageUrl: nextHref, loading: false, totalItems: totalItems, pageCache: newCache
      };
    });
  }
  // --- HELPER METHODS FOR QUERY BUILDING ---
  private buildFilterXml_Strict(): string {
    if (!this.state.filters) return '';
    const clauses: string[] = [];
    Object.keys(this.state.filters).forEach(key => {
      const f = this.state.filters[key];
      const val = (f && typeof f === 'object') ? f.value : f;
      if (val) clauses.push(`<Eq><FieldRef Name="${key}"/><Value Type="Text">${val}</Value></Eq>`);
    });
    if (clauses.length > 0) {
      let c = clauses[0];
      for (let i = 1; i < clauses.length; i++) c = `<And>${c}${clauses[i]}</And>`;
      return c;
    }
    return '';
  }
  private buildFilterXml_Loose(): string {
    if (!this.state.filters) return '';
    const clauses: string[] = [];
    Object.keys(this.state.filters).forEach(key => {
      const f = this.state.filters[key];
      const val = (f && typeof f === 'object') ? f.value : f;
      // Use Contains for flexibility in CAML mode
      if (val) clauses.push(`<Contains><FieldRef Name="${key}"/><Value Type="Text">${val}</Value></Contains>`);
    });
    if (clauses.length > 0) {
      let c = clauses[0];
      for (let i = 1; i < clauses.length; i++) c = `<And>${c}${clauses[i]}</And>`;
      return c;
    }
    return '';
  }
  private renderPagination_Above5K(): React.ReactElement<any> {
    const { page, nextPageUrl, loading, totalItems, pageSize } = this.state;
    //  If we have a Next Page URL, but Total Items equals the page size, 
    // it means we don't know the real total. Show a "+" sign.
    const showPlus = nextPageUrl && (totalItems <= (page * pageSize));
    return (
      <div className={styles.paginationFooter}>
        <div className={styles.paginationInfo}>
          {/* Shows "100+" if there are more pages but SharePoint didn't give a count */}
          Total Items: {totalItems}{showPlus ? '+' : ''} | Current Page: {page}
        </div>
        <div className={styles.paginationControls}>
          <button
            className={styles.btn}
            disabled={page === 1 || loading}
            onClick={this.onPrevPage_Large}
          >
            &lt; Previous
          </button>
          <button
            className={styles.btn}
            disabled={!nextPageUrl || loading}
            onClick={this.onNextPage_Large}
          >
            Next &gt;
          </button>
        </div>
        <div className={styles.perPage}>
          <label>Rows per page:</label>
          <select
            value={pageSize}
            disabled={loading}
            onChange={(e) => this.onPageSizeChange_Large(parseInt(e.target.value, 10))}
          >
            <option value="10">10</option>
            <option value="30">30</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>
      </div>
    );
  }
  private onNextPage_Large = (): void => {
    if (!this.state.nextPageUrl) return;
    //  Update Page Number
    const newPage = this.state.page + 1;
    this.setState({ page: newPage }, () => {
      //  Load with 'page' mode (Replaces items, uses NextHref)
      void this.loadItems_above5K('page');
    });
  }
  private onPrevPage_Large = (): void => {
    if (this.state.page <= 1) return;
    const prevPage = this.state.page - 1;
    //  Check Cache
    if (this.state.pageCache[prevPage]) {
      const cached = this.state.pageCache[prevPage];
      this.setState({
        page: prevPage,
        items: cached.items,
        nextPageUrl: cached.nextHref ?? undefined, // Important: Restore the "Next" token for this page
        loading: false
      });
    } else {
      // Fallback: If no cache, reset to start (Safest option for Large Lists)
      this.setState({ page: 1 }, () => void this.loadItems_above5K('init'));
    }
  }
  private onPageSizeChange_Large = (newSize: number): void => {
    this.setState({
      pageSize: newSize,
      page: 1,
      pageCache: {}, // Clear cache as page chunks have changed
      nextPageUrl: undefined
    }, () => {
      void this.loadItems_above5K('init');
    });
  }

  /**
   * Generates a link and copies it (Handles HTTP and HTTPS)
   */
  private handleShare(itemId: number): void {
    try {
      if (!itemId) return;
      const baseUrl = window.location.protocol + '//' + window.location.host + window.location.pathname;
      const shareUrl = `${baseUrl}?qmode=view&itemId=${itemId}`;
      //  Cast to 'any' to suppress the TypeScript error
      const nav = navigator as any;
      // Check if modern API is available AND we are on HTTPS
      if (nav.clipboard && window.isSecureContext) {
        nav.clipboard.writeText(shareUrl).then(() => {
          void Swal.fire({ icon: 'info', title: 'Info', text: 'Link copied to clipboard!' });
        }).catch((error: any) => {
          void LoggerService.log(
            'PowerForm - handleShare-Clipboard API failed',
            'Medium',
            itemId.toString(),
            error.message || JSON.stringify(error)
          );
          this.fallbackCopyTextToClipboard(shareUrl);
        });
      } else {
        // Fallback for HTTP environments (like your VM) or IE11
        this.fallbackCopyTextToClipboard(shareUrl);
      }
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleFilterChange',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  /**
   * Fallback for older browsers or HTTP sites
   */
  private fallbackCopyTextToClipboard(text: string) {
    const textArea = document.createElement("textarea");
    textArea.value = text;
    // Avoid scrolling to bottom
    textArea.style.top = "0";
    textArea.style.left = "0";
    textArea.style.position = "fixed";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    try {
      const successful = document.execCommand('copy');
      if (successful) {
        void Swal.fire({ icon: 'info', title: 'Info', text: 'Link copied to clipboard!' });
      } else {
        prompt('Copy this link manually:', text);
      }
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - fallbackCopyTextToClipboard',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      prompt('Copy this link manually:', text);
    }
    document.body.removeChild(textArea);
  }
  private applyClientSideFilters(): void {
    let items = this.allSourceItems || [];
    // =========================================================
    //  URL PARAMETER FILTERING (e.g. ?FilterField1=Status&FilterValue1=Active)
    // =========================================================
    try {
      const getQueryParam = (name: string): string | null => {
        const match = RegExp('[?&]' + name + '=([^&]*)').exec(window.location.search);
        return match && decodeURIComponent(match[1].replace(/\+/g, ' '));
      };
      for (let i = 1; i <= 10; i++) {
        const fFieldParam = getQueryParam('FilterField' + i);
        const fValue = getQueryParam('FilterValue' + i);
        if (fFieldParam && fValue) {
          let realKey = fFieldParam;
          if (this.state.fields) {
            const fieldDef = this.state.fields.find(f => f.InternalName === fFieldParam || f.EntityPropertyName === fFieldParam);
            if (fieldDef) {
              realKey = fieldDef.EntityPropertyName;
            }
          }
          items = items.filter(item => {
            const rawVal = item[realKey];
            if (rawVal === null || rawVal === undefined) return false;
            // Handle Objects (Lookups, Users)
            if (typeof rawVal === 'object') {
              const valStr = (rawVal.Title || rawVal.Name || rawVal.Id || '').toString().toLowerCase();
              return valStr === fValue.toLowerCase();
            }
            // Handle Arrays
            if (Array.isArray(rawVal)) {
              return rawVal.some((v: any) => {
                const subVal = typeof v === 'object' ? (v.Title || v.Name || '') : String(v);
                return subVal.toLowerCase() === fValue.toLowerCase();
              });
            }
            // Handle Simple Values
            return String(rawVal).toLowerCase() === fValue.toLowerCase();
          });
        }
      }
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - applyClientSideFilters',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
    // =========================================================
    //  GLOBAL SEARCH TEXT (Top Search Bar)
    // =========================================================
    if (this.state.searchText) {
      const search = this.state.searchText.toLowerCase();
      items = items.filter(item => {
        // Simple JSON stringify to search all properties
        return JSON.stringify(item).toLowerCase().indexOf(search) > -1;
      });
    }
    // =========================================================
    //  COLUMN LEVEL FILTERS (Inputs under headers)  
    // =========================================================
    const activeFilters = this.state.filters;
    if (activeFilters && Object.keys(activeFilters).length > 0) {
      Object.keys(activeFilters).forEach(key => {
        // Handle case where filter might be a string or an object { value: '...' }
        let realKey = key;
        if (this.state.fields) {
          // Find the field where InternalName matches the filter key
          const fieldDef = this.state.fields.find(f => f.InternalName === key);
          if (fieldDef) {
            realKey = fieldDef.EntityPropertyName;
          }
        }
        const rawFilter = activeFilters[key];
        const filterText = (rawFilter && typeof rawFilter === 'object' && (rawFilter as any).value)
          ? (rawFilter as any).value
          : String(rawFilter || '');
        if (filterText) {
          const lowerFilter = filterText.toLowerCase();
          items = items.filter(item => {
            const rawVal = item[realKey];
            // A. Handle Nulls
            if (rawVal === null || rawVal === undefined) return false;
            // B. Handle Arrays (Multi-Choice, Multi-User)
            if (Array.isArray(rawVal) || (rawVal.results && Array.isArray(rawVal.results))) {
              const arr = Array.isArray(rawVal) ? rawVal : rawVal.results;
              // Return true if ANY item in the array contains the filter text
              return arr.some((v: any) => {
                const subVal = (typeof v === 'object') ? (v.Title || v.Name || '') : String(v);
                return subVal.toLowerCase().indexOf(lowerFilter) > -1;
              });
            }
            // C. Handle Objects (Lookup, Single User, URL)
            if (typeof rawVal === 'object') {
              const valStr = (rawVal.Title || rawVal.Name || rawVal.Description || rawVal.Url || '').toString();
              return valStr.toLowerCase().indexOf(lowerFilter) > -1;
            }
            // D. Handle Simple Values (Text, Number, Date)
            // We use 'indexOf' for partial matching (e.g. typing "Inv" matches "Invoice")
            return String(rawVal).toLowerCase().indexOf(lowerFilter) > -1;
          });
        }
      });
    }
    // =========================================================
    //  NEW: CLIENT-SIDE SORTING LOGIC
    // =========================================================
    const { sortField, sortDirection, fields } = this.state;
    // Check if sortField exists and is valid in current fields array
    if (sortField && fields && fields.length > 0) {
      items.sort((a, b) => {
        const valA = this.getSafeStringValue(a, sortField).toLowerCase();
        const valB = this.getSafeStringValue(b, sortField).toLowerCase();
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);
        if (!isNaN(numA) && !isNaN(numB)) {
          return sortDirection === 'asc' ? numA - numB : numB - numA;
        }
        if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
        if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
        return 0;
      });
    }
    // =========================================================
    //  UPDATE STATE
    // ========================================================= 
    this.setState({
      items: items, // Update the main display array
      loading: false,
      page: 1 // Reset to page 1 when filtering
    });
  }
  // Helper to extract text from any field type (Note, User, Multi, etc.)
  private getSafeStringValue(item: any, key: string): string {
    try {
      const val = item[key];
      if (val === null || val === undefined) return '';
      // Handle nested objects (User/Lookup)
      if (typeof val === 'object') {
        if (val.Title) return val.Title; // Single User/Lookup
        if (val.Url) return val.Description || val.Url; // URL
        if (Array.isArray(val)) return val.map((v: any) => v.Title || v).join(', '); // MultiChoice / MultiUser
        if (val.results) return val.results.map((v: any) => v.Title || v).join(', '); // Old OData style Multi
      }
      // Handle standard values
      return String(val);
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - getSafeStringValue',
        'Low',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return '';
    }
  }
  /**
   * Checks if the user has permission for the requested URL mode.
   * If not, shows an error and redirects to the main list.
   */
  /**
     * Checks if the user has permission for the requested URL mode.
     * If not, shows an error and redirects to the main list.
     */
  private validateUrlPermissions(mode: string, perms: { canAdd: boolean, canEdit: boolean, canView: boolean }): boolean {
    try {
      let violation = false;
      let action = '';
      // Check specific violations
      // 1. ADD MODE: Block if no permission OR Override is checked
      if (mode === 'add') {
        if (!perms.canAdd || this.props.overrideAdd) {
          violation = true;
          action = 'create new items';
        }
      }
      // 2. EDIT MODE: Block if no permission OR Override is checked
      else if (mode === 'edit') {
        if (!perms.canEdit || this.props.overrideEdit) {
          violation = true;
          action = 'edit items';
        }
      }
      // 3. VIEW MODE: Block if no permission
      else if (mode === 'view') {
        if (!perms.canView) {
          violation = true;
          action = 'view items';
        }
      }
      if (violation) {
        this.setState({ loading: false }); // Stop loading spinner
        let timerInterval: any;
        void Swal.fire({
          icon: 'error',
          title: 'Access Denied',
          html: `You do not have permission to <b>${action}</b>.<br/>Redirecting to home in <b></b> seconds...`,
          timer: 5000,
          timerProgressBar: true,
          allowOutsideClick: false,
          allowEscapeKey: false,
          didOpen: () => {
            Swal.showLoading();
            const b = Swal.getHtmlContainer()?.querySelector('b:nth-child(2)');
            if (b) b.textContent = Math.ceil((Swal.getTimerLeft() ?? 0) / 1000).toString();
          },
          willClose: () => {
            clearInterval(timerInterval);
          }
        }).then((result) => {
          // Redirect logic
          this.redirectToHome();
        });
        return false; // Stop further execution
      }
      return true; // Access granted
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - validateUrlPermissions',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return false; // Default to denying access on error for safety
    }
  }
  /**
   * Clears query parameters to return to the default "List" view
   */
  private redirectToHome() {
    //  Construct URL without query params
    const cleanUrl = window.location.protocol + "//" + window.location.host + window.location.pathname;
    //  Refresh/Navigate
    window.location.href = cleanUrl;
  }

  private async loadItemData(id: number): Promise<void> {
    if (!id) return;
    this.setState({ loading: true, message: '' });

    try {
      const web = this._sp.web;

      // 1. LOAD PARENT ITEM
      const item = await web.lists.getByTitle(this.props.selectedList).items.getById(id)
        .select("*", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail", "Created", "Modified")
        .expand("Author", "Editor")
        ();

      const atts = await web.lists.getByTitle(this.props.selectedList).items.getById(id).attachmentFiles.select('FileName', 'ServerRelativeUrl')();

      // 2. PROCESS PEOPLE PICKER CACHE
      const personFields = this.state.fields.filter(f => f.TypeAsString === 'User' || f.TypeAsString === 'UserMulti');
      const peopleCache: { [key: number]: any } = {};

      for (let i = 0; i < personFields.length; i++) {
        const f = personFields[i];
        const idProp = f.InternalName + 'Id';

        // Collect all IDs
        let idsToFetch: number[] = [];
        if (f.AllowMultipleValues) {
          idsToFetch = item[idProp] && item[idProp].results ? item[idProp].results : [];
        } else {
          const singleId = item[idProp];
          if (singleId) idsToFetch.push(singleId);
        }

        // Fetch Data for each ID
        for (let j = 0; j < idsToFetch.length; j++) {
          const uid = idsToFetch[j];

          if (!peopleCache[uid]) {
            // --- START: ROBUST FETCH LOGIC ---
            let userResult: any = null;
            let principalType = 1; // Default to User

            try {
              // DEBUG LOG: Start
              console.log(`[PowerForm] Fetching ID ${uid} via getUserById...`);

              // ATTEMPT 1: Try getUserById (Standard)
              // We removed 'PrincipalType' from select to reduce 500 errors, we can infer it or get it from fallback
              userResult = await web.getUserById(uid).select('Id', 'Title', 'Email', 'LoginName', 'PrincipalType')();
              principalType = userResult.PrincipalType;

            } catch (errUser) {
              console.warn(`[PowerForm] getUserById(${uid}) failed (500/404). trying siteGroups...`, errUser);

              try {
                // ATTEMPT 2: Try Site Groups (Specific for SP Groups)
                userResult = await web.siteGroups.getById(uid).select('Id', 'Title', 'LoginName')();
                principalType = 8; // It is definitely a SharePoint Group
                console.log(`[PowerForm] ID ${uid} retrieved via siteGroups.`);

              } catch (errGroup) {
                console.warn(`[PowerForm] siteGroups.getById(${uid}) failed. trying UserInfoList...`, errGroup);

                try {
                  // ATTEMPT 3: User Info List (Ultimate Fallback - Reads raw list item)
                  // enable 'siteUserInfoList' property TS Error -> Use Root Web & GetByTitle
                  // We access the Site Root because User Info List lives there.

                  // We access the Site Root because User Info List lives there.
                  const rootWeb = Web([this._sp.web, this.props.context.pageContext.site.absoluteUrl]);

                  // Note: "User Information List" is the standard title. // If your site is in a different language, you might need to use the localized title or ID.
                  const userInfoItem = await rootWeb.lists.getByTitle("User Information List").items.getById(uid)
                    .select('Id', 'Title', 'Name', 'EMail', 'ContentTypeId')
                    ();

                  userResult = {
                    Id: userInfoItem.Id,
                    Title: userInfoItem.Title,
                    LoginName: userInfoItem.Name,
                    Email: userInfoItem.EMail
                  };

                  // Infer type roughly
                  principalType = (userInfoItem.ContentTypeId && userInfoItem.ContentTypeId.startsWith('0x010B')) ? 8 : 1;
                  console.log(`[PowerForm] ID ${uid} retrieved via UserInfoList (Root Web).`);

                } catch (errFallback) {
                  console.error(`[PowerForm] CRITICAL: Could not fetch principal ${uid} via any method.`, errFallback);
                }
              }
            }

            // MAPPING
            if (userResult) {
              // Logic to determine secondary text based on Type
              let subText = userResult.Email || "";
              if (!subText) {
                if (principalType === 8) subText = "SharePoint Group";
                else if (principalType === 4) subText = "Security Group";
                else subText = userResult.LoginName;
              }

              peopleCache[userResult.Id] = {
                key: userResult.Id.toString(),
                text: userResult.Title || "Unknown",
                primaryText: userResult.Title || "Unknown",
                secondaryText: subText,
                tertiaryText: userResult.LoginName,
                imageUrl: null,
                imageInitials: userResult.Title ? userResult.Title.split(' ').map((n: string) => n[0]).join('').slice(0, 2).toUpperCase() : '?'
              };
            }
            // --- END: ROBUST FETCH LOGIC ---
          }
        }
      }

      //todo check
      //this.setState({ formData: this.normaliseItemForForm(item), existingAttachments: atts, peopleOptions: peopleCache, attachments: atts.map((a: any) => ({ id: a.FileName, name: a.FileName, url: a.ServerRelativeUrl })), loading: false });


      // --- 3. LOAD CHILD ITEMS WITH LOGGING ---
      const loadedChildren: { [key: string]: any[] } = {};
      if (this.props.childConfigs && this.props.childConfigs.length > 0) {
        await Promise.all(this.props.childConfigs.map(async (conf) => {
          try {
            let filterKey = conf.foreignKeyField.trim();
            if (filterKey.indexOf('Id') === -1) { filterKey += 'Id'; }


            const childRows = await web.lists.getByTitle(conf.childListTitle).items
              .filter(`${filterKey} eq ${id}`)
              .top(500)
              ();
            loadedChildren[conf.childListTitle] = childRows;
          } catch (e: any) {
            console.error(`[PowerForm] Failed to load child items for ${conf.childListTitle}`, e);
          }
        }));
      } else {

      }

      // 4. SET STATE
      const formData = this.normaliseItemForForm(item);


      this.setState({
        formData: formData,
        existingAttachments: atts,
        peopleOptions: peopleCache,
        attachments: atts.map((a: any) => ({ id: a.FileName, name: a.FileName, url: a.ServerRelativeUrl })),
        childItems: loadedChildren, // <--- UPDATE CHILD ITEMS STATE
        loading: false
      }, () => {
        // Trigger Cascade
        if (this.props.cascadeConfig) {
          Object.keys(this.props.cascadeConfig).forEach(childKey => {
            const conf = this.props.cascadeConfig![childKey];
            const parentVal = this.state.formData[conf.parentField];
            if (parentVal) {
              void this.loadCascadeOptions(childKey, parentVal);
            }
          });
        }
      });

      if (this.props.enableNotification && this.state.mode === 'view') {

        const config = this.getNotificationConfig();
        const listName = this.props.listPageTitle || this.props.selectedList || "List";

        void NotificationService.logNotification(
          this.service.siteUrl,
          listName,
          'Viewed',
          item,
          this.state.currentUser,
          config
        );

      }

    } catch (error: any) {
      console.error("[PowerForm] loadItemData Error:", error);
      this.setState({ loading: false });
    }
  }

  private renderChildSection(config: IChildListConfig): JSX.Element {
    // 1. Get items from state
    const items = this.state.childItems ? (this.state.childItems[config.childListTitle] || []) : [];
    const itemsJson = JSON.stringify(items);

    // 2. Build Columns
    const columns: IRepeaterColumn[] = config.visibleFields.map(f => ({
      key: f,
      name: f,
      type: 'text' as const,
      required: false
    }));

    return (
      <div key={config.childListTitle} className={styles.formSection} style={{ marginTop: '20px', padding: '10px', border: '1px solid #eaeaea', background: '#f9f9f9' }}>
        <h3 style={{ margin: '0 0 10px 0', fontSize: '15px', color: '#0078d4' }}>{config.title || config.childListTitle}</h3>

        <RepeaterInput
          mode={this.state.mode === 'view' ? 'view' : 'edit'}
          columns={columns}
          value={itemsJson}
          onChange={(newVal) => {
            try {
              const updatedItems = JSON.parse(newVal);
              this.setState(prevState => ({
                childItems: {
                  ...prevState.childItems,
                  [config.childListTitle]: updatedItems
                }
              }));
            } catch (e: any) { console.error("Error updating child state:", e); }
          }}
        />
      </div>
    );
  }


  //  validateField: Merges errors and cleans up empty keys
  private validateField = async (field: ColumnDefinition, value: any): Promise<void> => {
    try {
      // 1. Standard Validation (Required, etc.)
      let error = this.getStandardValidation(field, value);

      // 2. Unique Value Check (Async)
      // Only run if standard validation passed and the field requires uniqueness
      if (!error && field.EnforceUniqueValues) {
        const uniqueError = await this.checkIfValueIsUnique(field, value);
        // eslint-disable-next-line require-atomic-updates
        error = uniqueError || '';
      }

      // 3. Custom Validation (Business Rules)
      // Run in 'blur' mode so specific onBlur rules trigger
      if (!error) {
        // eslint-disable-next-line require-atomic-updates
        error = this.runCustomValidations(field, value, 'blur', this.state.formData);
      }

      // 4. SAFE STATE UPDATE
      this.setState((prevState) => {
        // Create a shallow copy of the existing errors to avoid mutation
        const mergedErrors = { ...prevState.formErrors };

        if (error) {
          // If there is an error, update/add it
          mergedErrors[field.EntityPropertyName] = error;
        } else {
          // [CRITICAL] If error is cleared, DELETE the key.
          // This ensures Object.keys(formErrors).length returns 0 when valid.
          delete mergedErrors[field.EntityPropertyName];
        }

        return { formErrors: mergedErrors };
      });

    } catch (err: any) {
      void LoggerService.log(
        'PowerForm - validateField',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        err.message || JSON.stringify(err)
      );
    }
  }
  private checkIfValueIsUnique = async (field: ColumnDefinition, value: string): Promise<string> => {
    if (!value || !value.toString().trim()) return '';
    const fieldInternalName = field.EntityPropertyName;
    const listName = this.props.selectedList;
    try {
      // PnPjs: Fetch items with the same value (Limit to 5 to be safe)
      const safeValue = String(value).replace(/'/g, "''"); // Escape single quotes for OData
      const items = await this._sp.web.lists.getByTitle(listName).items
        .filter(`${fieldInternalName} eq '${safeValue}'`)
        .select('Id')
        .top(5)();
        
      // FILTER: Exclude the current item if we are in Edit Mode
      const duplicates = items.filter((item: any) => {
        // If editing, and this item's ID matches the current Item ID, it's NOT a duplicate (it's the same item)
        if (this.state.mode === 'edit' && this.state.itemId && item.Id === this.state.itemId) {
          return false;
        }
        return true;
      });
      if (duplicates.length > 0) {
        return `${field.Title} must be unique. "${value}" already exists.`;
      }
      return '';
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - checkIfValueIsUnique',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return '';
    }
  }
  // Helper to check if the form is valid (No Errors + All Required Fields Filled)
  private isFormInvalid(visibleFields: ColumnDefinition[]): boolean {
    try {
      //  Check for Active Errors in State (Min/Max, Unique, etc.)
      const hasErrors = Object.keys(this.state.formErrors).some(key => {
        const err = this.state.formErrors[key];
        return err && err.length > 0;
      });
      if (hasErrors) return true;
      //  Check for Empty Required Fields
      // We only check fields that are currently visible on the form
      for (const f of visibleFields) {
        if (f.Required) {
          const val = this.state.formData[f.EntityPropertyName];
          // Check for Empty Strings, Null, Undefined
          if (val === null || val === undefined || val === '') return true;
          // Check for Empty Arrays (MultiChoice, User, Lookup)
          if (Array.isArray(val) && val.length === 0) return true;
        }
      }
      return false; // Form is VALID
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - isFormInvalid',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return true; // Fail safe: block submission if validation crashes
    }
  }
  private runCustomValidations(field: ColumnDefinition, value: any, trigger: string, currentFormData?: any): string {
    try {
      const config = this.props.validationConfig || {};
      const rules = config[field.InternalName] || [];
      const dataContext = currentFormData || this.state.formData;
      for (let r of rules) {
        if (trigger && r.trigger !== trigger && !(trigger === 'change' && r.trigger === 'blur')) {
          continue;
        }
        switch (r.type) {
          case 'regex': {
            try {
              if (r.pattern && !new RegExp(r.pattern).test(String(value || ''))) return r.message;
            } catch (regErr) {
              void LoggerService.log('PowerForm - runCustomValidations', 'Medium', 'N/A', `Invalid Regex Pattern`);
            }
            break;
          }
          case 'range': {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              if (r.min != null && num < Number(r.min)) return r.message;
              if (r.max != null && num > Number(r.max)) return r.message;
            }
            break;
          }
          case 'compare': {
            if (r.otherField && r.operator) {
              let otherKey = r.otherField;
              if (this.state.fields) {
                const otherFieldDef = this.state.fields.find(f => f.InternalName === r.otherField);
                if (otherFieldDef) otherKey = otherFieldDef.EntityPropertyName;
              }
              const otherVal = dataContext[otherKey];
              const isEmpty = (v: any) => v === null || v === undefined || String(v).trim() === '';
              if (!isEmpty(value) && !isEmpty(otherVal)) {
                let v1: any = value;
                let v2: any = otherVal;
                const n1 = Number(v1);
                const n2 = Number(v2);
                if (!isNaN(n1) && !isNaN(n2) && String(v1).trim() !== '' && String(v2).trim() !== '') {
                  v1 = n1; v2 = n2;
                } else {
                  const d1 = Date.parse(v1); const d2 = Date.parse(v2);
                  if (!isNaN(d1) && !isNaN(d2)) { v1 = d1; v2 = d2; }
                }
                let isValid = true;
                switch (r.operator) {
                  case 'eq': isValid = v1 == v2; break;
                  case 'ne': isValid = v1 != v2; break;
                  case 'gt': isValid = v1 > v2; break;
                  case 'ge': isValid = v1 >= v2; break;
                  case 'lt': isValid = v1 < v2; break;
                  case 'le': isValid = v1 <= v2; break;
                }
                if (!isValid) return r.message;
              }
            }
            break;
          }
          case 'custom': {
            if (r.fnBody) {
              try {
                const FuncBuilder = (window as any).Function;
                const fn = new FuncBuilder('value', 'formData', r.fnBody);
                if (!fn(value, dataContext)) return r.message;
              } catch (error: any) {
                void LoggerService.log('PowerForm - runCustomValidations', 'Medium', 'N/A', error.message);
              }
            }
            break;
          }
        }
      }
      return '';
    } catch (error: any) {
      void LoggerService.log('PowerForm - runCustomValidations', 'Medium', 'N/A', error.message);
      return '';
    }
  }
  private formatDate(dateString: string): string {
    try {
      if (!dateString) return '-';
      const date = new Date(dateString);
      const now = new Date();
      const diffMs = now.getTime() - date.getTime();
      const diffSec = Math.floor(diffMs / 1000);
      const diffMin = Math.floor(diffSec / 60);
      const diffHour = Math.floor(diffMin / 60);
      const diffDays = Math.floor(diffHour / 24);
      if (diffSec < 60) return 'Just now';
      if (diffMin < 60) return `${diffMin} minute${diffMin !== 1 ? 's' : ''} ago`;
      if (diffHour < 24) return `${diffHour} hour${diffHour !== 1 ? 's' : ''} ago`;
      if (diffDays === 1) return 'Yesterday';
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const d = date.getDate();
      const m = months[date.getMonth()];
      const y = date.getFullYear();
      let h = date.getHours();
      const min = date.getMinutes();
      const ampm = h >= 12 ? 'pm' : 'am';
      h = h % 12;
      h = h ? h : 12;
      const minStr = min < 10 ? '0' + min : min;
      return `${d} ${m} ${y} ${h}:${minStr} ${ampm}`;
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - formatDate',
        'Low',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return '-';
    }
  }
  private formatDate_form(dateString: string, isDateOnly: boolean = false): string {
    try {
      if (!dateString) return '-';
      // If Date Only, show local date string (no time)
      if (isDateOnly) return new Date(dateString).toLocaleDateString();
      // Otherwise show full date and time
      return new Date(dateString).toLocaleString();
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - formatDate_form',
        'Low',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return '-';
    }
  }
  private formatDateTimeForInput(dateString: string, isDateOnly: boolean = false): string {
    try {
      if (!dateString) return '';
      const date = new Date(dateString);
      const pad = (n: number) => (n < 10 ? '0' + n : n.toString());
      const y = date.getFullYear();
      const m = pad(date.getMonth() + 1);
      const d = pad(date.getDate());
      if (isDateOnly) {
        // Return YYYY-MM-DD for type="date"
        return `${y}-${m}-${d}`;
      }
      // Return YYYY-MM-DDTHH:mm for type="datetime-local"
      return `${y}-${m}-${d}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - formatDateTimeForInput',
        'Low',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return '';
    }
  }
  private getValidationMessage = (field: ColumnDefinition, value: any): string => {
    try {
      let msg = '';
      if (field.Required && (value === null || value === undefined || value === '')) msg = `${field.Title} is required.`;
      if (!msg) msg = this.runCustomValidations(field, value, 'blur');
      this.setState(prev => ({ formErrors: { ...prev.formErrors, [field.EntityPropertyName]: msg } }));
      return msg;
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - getValidationMessage',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return ''; // Return empty string so execution can continue safely
    }
  }
  private handleViewChange = (viewId: string) => {
    try {
      //  Reset if "All Items" selected
      if (!viewId) {
        this.setState({
          currentViewId: '',
          activeViewFields: null, // Reset to default list columns
          filters: {},
          page: 1
        }, () => void this.loadItems('init'));
        return;
      }
      //  Find the selected View config
      const view = this.state.availableViews.find(v => v.id === viewId);
      if (!view) return;
      //  Prepare Filters from the View
      let newFilters: any = {};
      if (view.filters) {
        view.filters.forEach((f: any) => {
          if (f.field && f.value) {
            newFilters[f.field] = {
              value: f.value,
              operator: f.operator
            };
          }
        });
      }
      //  Update State & Reload
      this.setState({
        currentViewId: viewId,
        activeViewFields: view.visibleFields || null, // <--- APPLY COLUMNS
        filters: newFilters,
        page: 1
      }, () => {
        void this.loadItems('filter');
      });
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleViewChange',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  // UPDATE: Accept fullItemData (optional) to handle column mapping
  private handleChange_hold = (fieldName: string, value: any, fullItemData?: any): void => {
    try {
      //  Standard Update
      const newFormData = { ...this.state.formData, [fieldName]: value };
      const newFormErrors = { ...this.state.formErrors };
      //  COLUMN MAPPING LOGIC (Auto-Populate)
      if (fullItemData) {
        // Helper to apply mappings
        const applyMap = (mappings: any[], type: string) => {
          mappings.forEach(map => {
            if (map.source && map.target) {
              // Try getting value
              const val = fullItemData[map.source] || fullItemData[this.toODataName(map.source)];
              // Find Target Field
              const targetFieldDef = this.state.fields.filter(function (f) {
                return f.InternalName === map.target;
              })[0];
              if (targetFieldDef) {
                const targetKey = targetFieldDef.EntityPropertyName;
                newFormData[targetKey] = val;
                // Clear error if valid
                if (newFormErrors[targetKey]) {
                  delete newFormErrors[targetKey];
                }
              } else {
                void LoggerService.log(
                  'PowerForm - handleChange',
                  'Medium',
                  this.state.itemId ? this.state.itemId.toString() : 'N/A',
                  `[Mapping] Target field '${map.target}' not found in current list fields.`
                );
              }
            }
          });
        };
        // A. Check Autocomplete Config
        let acConfig = this.props.autocompleteConfig ? this.props.autocompleteConfig[fieldName] : null;
        if (!acConfig && this.props.autocompleteConfig && this.state.fields) {
          const fieldMatch = this.state.fields.find(function (f) { return f.EntityPropertyName === fieldName; });
          if (fieldMatch) {
            acConfig = this.props.autocompleteConfig[fieldMatch.InternalName];
          }
        }
        if (acConfig && acConfig.columnMapping) {
          applyMap(acConfig.columnMapping, 'Autocomplete');
        }
        // B. Check Lookup Config
        let luConfig = this.props.lookupDisplayConfig ? this.props.lookupDisplayConfig[fieldName] : null;
        if (!luConfig && this.state.fields) {
          const fieldMatch = this.state.fields.find(function (f) { return f.EntityPropertyName === fieldName; });
          const fInternal = fieldMatch ? fieldMatch.InternalName : null;
          if (fInternal && this.props.lookupDisplayConfig) {
            luConfig = this.props.lookupDisplayConfig[fInternal];
          }
        }
        if (luConfig && luConfig.columnMapping) {
          applyMap(luConfig.columnMapping, 'Lookup');
        }
        // C. Check Cascade Config
        let casConfig = this.props.cascadeConfig ? this.props.cascadeConfig[fieldName] : null;
        if (!casConfig && this.state.fields) {
          const fieldMatch = this.state.fields.find(function (f) { return f.EntityPropertyName === fieldName; });
          const fInternal = fieldMatch ? fieldMatch.InternalName : null;
          if (fInternal && this.props.cascadeConfig) {
            casConfig = this.props.cascadeConfig[fInternal];
          }
        }
        if (casConfig && casConfig.columnMapping) {
          applyMap(casConfig.columnMapping, 'Cascade');
        }
      }
      //  Validate CURRENT Field
      const fieldMeta = this.state.fields.find(function (f) { return f.EntityPropertyName === fieldName; });
      if (fieldMeta) {
        let error = this.getStandardValidation(fieldMeta, value);
        if (!error) {
          error = this.runCustomValidations(fieldMeta, value, 'change', newFormData);
        }
        newFormErrors[fieldName] = error;
      }
      //  Validate DEPENDENT Fields
      const config = this.props.validationConfig || {};
      Object.keys(config).forEach(otherInternalName => {
        if (fieldMeta && otherInternalName === fieldMeta.InternalName) return;
        const rules = config[otherInternalName];
      const dependsOnCurrent = fieldMeta ? rules.some(r => r.type === 'compare' && r.otherField === fieldMeta.InternalName) : false;
        if (dependsOnCurrent) {
          const otherFieldDef = this.state.fields.find(function (f) { return f.InternalName === otherInternalName; });
          if (otherFieldDef) {
            const otherKey = otherFieldDef.EntityPropertyName;
            const otherValue = newFormData[otherKey];
            const depError = this.runCustomValidations(otherFieldDef, otherValue, 'change', newFormData);
            newFormErrors[otherKey] = depError;
          }
        }
      });
      //  Update State
      this.setState({ formData: newFormData, formErrors: newFormErrors });
      //  Cascade Logic
      if (this.props.cascadeConfig) {
        Object.keys(this.props.cascadeConfig).forEach(childKey => {
          const conf = this.props.cascadeConfig![childKey];
          if (conf.parentField === fieldName) {
            this.setState(prev => ({
              formData: { ...prev.formData, [childKey]: null }
            }));
            void this.loadCascadeOptions(childKey, value);
          }
        });
      }
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleChange',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private _resolvePeopleSuggestions = async (filter: string): Promise<IPersonaProps[]> => {
    if (!filter) return [];
    try {
     // Modern PnPjs People Search API
      const results = await this._sp.profiles.clientPeoplePickerSearchUser({
        QueryString: filter,
        MaximumEntitySuggestions: 5,
        AllowEmailAddresses: true,
        PrincipalSource: 15, // All Sources
        PrincipalType: 15    // Users + SecGroups + SPGroups
      });

      const suggestions = await Promise.all(results.map(async (r: any) => {
        let userId = (r.EntityData && r.EntityData.SPUserID) ? r.EntityData.SPUserID :
                     (r.EntityData && r.EntityData.SPGroupID) ? r.EntityData.SPGroupID : null;
                     
        // Ensure user exists in the site collection if we only have their Key
        if (!userId && r.Key) {
          try {
            const ensureRes = await this._sp.web.ensureUser(r.Key);
            if (ensureRes) userId = ensureRes.Id.toString();
          } catch (e: any) {
            return null;
          }
        }
        
        if (!userId) return null;
        
        return {
          key: userId,
          text: r.DisplayText,
          primaryText: r.DisplayText,
          secondaryText: (r.EntityData && r.EntityData.Email) ? r.EntityData.Email : (r.Description || r.EntityType),
          imageInitials: r.DisplayText ? r.DisplayText.split(' ').map((n: string) => n[0]).join('').substring(0, 2).toUpperCase() : '?'
        } as IPersonaProps;
      }));
      
      return suggestions.filter((s): s is IPersonaProps => s !== null);
    } catch (e: any) {
      void LoggerService.log('PowerForm - _resolvePeopleSuggestions', 'Medium', 'N/A', e.message);
      return [];
    }
  }
  //  handleChange: Safely merges errors and handles all logic
  private handleChange = (fieldName: string, value: any, fullItemData?: any): void => {
    try {
      // 1. Prepare Data Update (Input is authoritative)
      const newFormData = { ...this.state.formData, [fieldName]: value };

      // 2. Prepare Error Updates (We only track what CHANGED, to merge later)
      const errorsToUpdate: { [key: string]: string } = {};

      // --- SECTION A: COLUMN MAPPING LOGIC (Auto-Populate) ---
      if (fullItemData) {
        // Helper to apply mappings for Autocomplete/Lookup/Cascade
        const applyMap = (mappings: any[]) => {
          mappings.forEach(map => {
            if (map.source && map.target) {
              // Try getting value from source (check standard and OData names)
              const val = fullItemData[map.source] || fullItemData[this.toODataName(map.source)];

              // Find Target Field Definition
              const targetFieldDef = this.state.fields.find(f => f.InternalName === map.target);
              if (targetFieldDef) {
                const targetKey = targetFieldDef.EntityPropertyName;
                // Update Data
                newFormData[targetKey] = val;

                // [CRITICAL] Mark error as cleared for this mapped field (since it now has a value)
                errorsToUpdate[targetKey] = "";
              } else {
                void LoggerService.log(
                  'PowerForm - handleChange',
                  'Medium',
                  this.state.itemId ? this.state.itemId.toString() : 'N/A',
                  `[Mapping] Target field '${map.target}' not found in current list fields.`
                );
              }
            }
          });
        };

        // 1. Check Autocomplete Config
        let acConfig = this.props.autocompleteConfig ? this.props.autocompleteConfig[fieldName] : null;
        if (!acConfig && this.props.autocompleteConfig && this.state.fields) {
          // Fallback: Check by Internal Name if EntityPropertyName didn't match
          const fieldMatch = this.state.fields.find(f => f.EntityPropertyName === fieldName);
          if (fieldMatch) {
            acConfig = this.props.autocompleteConfig[fieldMatch.InternalName];
          }
        }
        if (acConfig && acConfig.columnMapping) {
          applyMap(acConfig.columnMapping);
        }

        // 2. Check Lookup Config
        let luConfig = this.props.lookupDisplayConfig ? this.props.lookupDisplayConfig[fieldName] : null;
        if (!luConfig && this.state.fields) {
          const fieldMatch = this.state.fields.find(f => f.EntityPropertyName === fieldName);
          const fInternal = fieldMatch ? fieldMatch.InternalName : null;
          if (fInternal && this.props.lookupDisplayConfig) {
            luConfig = this.props.lookupDisplayConfig[fInternal];
          }
        }
        if (luConfig && luConfig.columnMapping) {
          applyMap(luConfig.columnMapping);
        }

        // 3. Check Cascade Config
        let casConfig = this.props.cascadeConfig ? this.props.cascadeConfig[fieldName] : null;
        if (!casConfig && this.state.fields) {
          const fieldMatch = this.state.fields.find(f => f.EntityPropertyName === fieldName);
          const fInternal = fieldMatch ? fieldMatch.InternalName : null;
          if (fInternal && this.props.cascadeConfig) {
            casConfig = this.props.cascadeConfig[fInternal];
          }
        }
        if (casConfig && casConfig.columnMapping) {
          applyMap(casConfig.columnMapping);
        }
      }

      // --- SECTION B: VALIDATE CURRENT FIELD ---
      const fieldMeta = this.state.fields.find(f => f.EntityPropertyName === fieldName);
      if (fieldMeta) {
        // 1. Standard Validation (Required, etc.)
        let error = this.getStandardValidation(fieldMeta, value);

        // 2. Custom Validation (Regex, Range, etc.)
        if (!error) {
          // 'change' mode means we check sync rules. Async rules usually run on 'blur'.
          error = this.runCustomValidations(fieldMeta, value, 'change', newFormData);
        }

        // Add result to our update list
        errorsToUpdate[fieldName] = error;
      }

      // --- SECTION C: VALIDATE DEPENDENT FIELDS ---
      // (If Field B compares itself to Field A, we must re-validate Field B now)
      const config = this.props.validationConfig || {};
      Object.keys(config).forEach(otherInternalName => {
        // Skip self
        if (fieldMeta && otherInternalName === fieldMeta.InternalName) return;

        const rules = config[otherInternalName];
        // Check if any rule for 'otherField' depends on 'currentField'
     const dependsOnCurrent = fieldMeta ? rules.some(r => r.type === 'compare' && r.otherField === fieldMeta.InternalName) : false;
        if (dependsOnCurrent) {
          const otherFieldDef = this.state.fields.find(f => f.InternalName === otherInternalName);
          if (otherFieldDef) {
            const otherKey = otherFieldDef.EntityPropertyName;
            const otherValue = newFormData[otherKey];

            // Re-run validation for the dependent field
            const depError = this.runCustomValidations(otherFieldDef, otherValue, 'change', newFormData);
            errorsToUpdate[otherKey] = depError;
          }
        }
      });

      // --- SECTION D: SAFE STATE UPDATE  ---
      // We use a Functional State Update (prevState => ...) to ensure we merge
      // with the latest errors, specifically preserving errors set by onBlur/Async events.
      this.setState(prevState => {
        // 1. Merge previous errors with new updates
        const mergedErrors = { ...prevState.formErrors, ...errorsToUpdate };

        // 2. Cleanup: Remove keys where error is empty string to keep state clean
        Object.keys(errorsToUpdate).forEach(k => {
          if (errorsToUpdate[k] === "") delete mergedErrors[k];
        });

        return {
          formData: newFormData,
          formErrors: mergedErrors
        };
      });

      // --- SECTION E: HANDLE CASCADE DROPDOWNS ---
      // If this field is a Parent, reset the Child fields
      if (this.props.cascadeConfig) {
        Object.keys(this.props.cascadeConfig).forEach(childKey => {
          const conf = this.props.cascadeConfig![childKey];
          if (conf.parentField === fieldName) {
            // Reset child data in State
            this.setState(prev => ({
              formData: { ...prev.formData, [childKey]: null }
            }));
            // Trigger load for new options
            void this.loadCascadeOptions(childKey, value);
          }
        });
      }

    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleChange',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    try {
      if (e.target.files) {
        const files = [].slice.call(e.target.files) as File[];
        this.setState(prev => ({ attachmentsNew: [...prev.attachmentsNew, ...files] }));
      }
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleFileChange',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  // --- HELPER: Construct Clean Payload for Parent Item ---
  // --- HELPER: Construct Clean Payload for Parent Item ---
  private buildRestPayload(): any {
    const payload: any = {};
    const { fields, formData } = this.state;

    // Iterate through defined fields only
    fields.forEach(f => {
      const key = f.EntityPropertyName; // e.g. "AssignedTo"
      const val = formData[key];        // e.g. [12, 15] or 12

      // 1. SKIP System/ReadOnly Fields
      if (f.InternalName === 'ID' || f.InternalName === 'Author' || f.InternalName === 'Editor' ||
        f.InternalName === 'Created' || f.InternalName === 'Modified' || f.InternalName === 'Attachments' ||
        f.InternalName === 'ContentType' || (f as any).ReadOnlyField) {
        return;
      }

      // 2. Handle User Fields  
      if (f.TypeAsString === 'User' || f.TypeAsString === 'UserMulti') {
        const idKey = f.InternalName + 'Id'; // e.g. "AssignedToId"

        // We check 'val' (from EntityPropertyName) because that's where handleChange saves the data
        if (val !== undefined && val !== null) {
          if (f.TypeAsString === 'UserMulti') {
            // Multi-User must be wrapped in 'results' for SharePoint REST
            payload[idKey] = { results: Array.isArray(val) ? val : [val] };
          } else {
            // Single User is just the ID
            payload[idKey] = val;
          }
        }
      }
      // 3. Handle Lookup Fields (Same logic as User)
      else if (f.TypeAsString === 'Lookup' || f.TypeAsString === 'LookupMulti') {
        const idKey = f.InternalName + 'Id';
        if (val !== undefined && val !== null) {
          if (f.TypeAsString === 'LookupMulti') {
            payload[idKey] = { results: Array.isArray(val) ? val : [val] };
          } else {
            payload[idKey] = Number(val);
          }
        }
      }
      // 4. Handle Numbers
      else if (f.TypeAsString === 'Number' || f.TypeAsString === 'Currency') {
        if (val === '' || val === undefined || val === null) {
          payload[key] = null;
        } else {
          payload[key] = Number(val);
        }
      }
      // 5. Handle MultiChoice
      else if (f.TypeAsString === 'MultiChoice') {
        if (val) {
          // MultiChoice also typically needs 'results' wrapper in standard REST, 
          // though some endpoints accept raw arrays. Safer to wrap.
          payload[key] = { results: Array.isArray(val) ? val : [val] };
        }
      }
      // 6. Standard Assignment
      else {
        if (val !== undefined) {
          payload[key] = val;
        }
      }
    });

    return payload;
  }

  // Handler: Remove item from the child list buffer
  private removeChildItem = (listKey: string, index: number): void => {
    const currentItems = this.state.childItems[listKey] || [];
    const newItems = [...currentItems];
    newItems.splice(index, 1);

    this.setState(prev => ({
      childItems: { ...prev.childItems, [listKey]: newItems }
    }));
  }

  // Handler: Open the modal to add a child item (Form Mode)
  private openChildForm = (config: any): void => {
    this.setState({
      isChildPanelOpen: true,
      activeChildConfig: config,
      activeChildItemIndex: -1 // -1 indicates "New Item"
    });
  }


  private handleSubmit = async (e?: React.FormEvent<HTMLFormElement>): Promise<void> => {
    if (e) e.preventDefault();
    this.allSourceItems = [];

    // -------------------------------------------------------------------------
    // 1. VALIDATION
    // -------------------------------------------------------------------------
    const errors: string[] = [];
    for (let i = 0; i < this.state.fields.length; i++) {
      const f = this.state.fields[i];
      const val = this.state.formData[f.EntityPropertyName];
      const err = this.getValidationMessage(f, val);
      if (err) errors.push(`${f.Title}: ${err}`);
    }

    if (errors.length > 0) {
      this.setState({ message: 'Validation failed.' });
      void Swal.fire({ icon: 'warning', title: 'Validation Error', html: errors.join('<br/>') });
      return;
    }

    // -------------------------------------------------------------------------
    // 2. SAVE PROCESS
    // -------------------------------------------------------------------------
    void Swal.fire({
      title: 'Saving...',
      html: 'Processing items...',
      allowOutsideClick: false,
      didOpen: () => { Swal.showLoading(); }
    });

    try {
      const web = this._sp.web;
      const list = web.lists.getByTitle(this.props.selectedList);

      // PRE-FETCH USER FOR NOTIFICATIONS (Efficiency)


      // Build Payload
      const payload = this.buildRestPayload();

      let parentId = this.state.itemId;

      //TODO : Check logic
      // DETERMINE PARENT ACTION & TITLE
      const parentAction = this.state.mode === 'add' ? 'Added' : 'Updated';
      let parentTitle = this.state.formData['Title'];

      // Fallback Title Logic
      if (!parentTitle) {
        const titleField = this.state.fields.find(f => f.InternalName === 'Title');
        if (titleField) parentTitle = this.state.formData[titleField.EntityPropertyName];
      }
      if (!parentTitle) parentTitle = `Item ${parentId || 'New'}`;

      // --- STEP A: SAVE PARENT ITEM ---
      if (this.state.mode === 'add') {
        const result = await list.items.add(payload);
        // eslint-disable-next-line require-atomic-updates
        parentId = result.data.Id;
        // eslint-disable-next-line require-atomic-updates
        if (result.data.Title) parentTitle = result.data.Title;
      } else {
        if (parentId) {
          const itemToUpdate = list.items.getById(parentId);
          await itemToUpdate.update(payload, "*");
        } else {
          throw new Error("Edit Mode but no Parent ID found.");
        }
      }

      // --- STEP B: SAVE ATTACHMENTS ---
      if (this.state.attachmentsNew && this.state.attachmentsNew.length > 0) {
        await Promise.all(this.state.attachmentsNew.map(file => {
          return list.items.getById(parentId ?? 0).attachmentFiles.add(file.name, file);
        }));
      }
      if (this.state.attachmentsDelete && this.state.attachmentsDelete.length > 0) {
        await Promise.all(this.state.attachmentsDelete.map(fileName => {
          return list.items.getById(parentId ?? 0).attachmentFiles.getByName(fileName).delete();
        }));
      }

      // --- NOTIFICATION 1: PARENT (Sent FIRST as requested) ---
      if (this.props.enableNotification && this.state.currentUser) {
        const config = this.getNotificationConfig();
        const listName = this.props.listPageTitle || this.props.selectedList || "List";

        const itemContext = {
          ...this.state.formData,
          Id: parentId,
          Title: parentTitle
        };

        // Log Parent
        void NotificationService.logNotification(
          this.service.siteUrl,
          listName, // Parent List Name
          parentAction,
          itemContext,
          this.state.currentUser,
          config
        ).catch(e => console.error("Parent Notification Error", e));
      }

      // --- STEP C: SAVE CHILD ITEMS & NOTIFY CHILDREN ---
      if (this.props.childConfigs && this.props.childConfigs.length > 0) {
        await Promise.all(this.props.childConfigs.map(async (conf) => {
          const childListKey = conf.childListTitle;
          const items = this.state.childItems[childListKey] || [];

          if (items.length > 0) {
            const childList = web.lists.getByTitle(childListKey);

            // Foreign Key Logic
            let fkField = conf.foreignKeyField;
            if (fkField.slice(-2) !== 'Id') {
              fkField += 'Id';
            }

            for (const row of items) {


              const childPayload: any = {};

              // Define System Fields to Ignore
              const ignoredKeys = [
                'ID', 'Id', 'Author', 'AuthorId', 'Editor', 'EditorId', 'Created', 'Modified',
                'GUID', 'AttachmentFiles', 'ContentType', 'ContentTypeId', 'ComplianceAssetId',
                'OData__UIVersionString', 'FileSystemObjectType', 'ServerRedirectedEmbedUri',
                'ServerRedirectedEmbedUrl', '__metadata'
              ];

              Object.keys(row).forEach(key => {
                // A. Skip ignored keys
                if (ignoredKeys.indexOf(key) > -1) return;

                // B. Skip OData fields
                if (key.indexOf('OData_') === 0) return;

                const val = row[key];

                // C. CRITICAL CHECK: Check for System Objects (The cause of your error)
                if (val && typeof val === 'object') {
                  // If it contains '__deferred' or '__metadata', it is a system object -> SKIP IT
                  if (val['__deferred'] || val['__metadata']) {
                    console.warn(`[DEBUG] Removed System Field: ${key}`, val);
                    return;
                  }
                }

                // D. Add valid data
                childPayload[key] = val;
              });

              // 1. Link to Parent
              childPayload[fkField] = parentId;

              // 3. Determine Child Action & Save
              let childId = row.Id;
              let childAction: 'Added' | 'Updated' = 'Added';

              try {
                if (childId && parseInt(childId) > 0) {
                  // UPDATE CHILD
                  childAction = 'Updated';
                  await childList.items.getById(childId).update(childPayload);
                } else {
                  childAction = 'Added';
                  const childResult = await childList.items.add(childPayload);
                  childId = childResult.data.Id;
                  // eslint-disable-next-line require-atomic-updates
                  if (childResult.data.Title) childPayload.Title = childResult.data.Title;
                }
              } catch (childErr) {
                console.error("[DEBUG] Child Save Failed!", childErr);
                throw childErr; // Throw so the main catch block handles the alert
              }

              // --- NOTIFICATION 2: CHILD ---
              if (this.props.enableNotification) {
                const config = this.getNotificationConfig();
                const childTitle = childPayload.Title || row.Title || `Item ${childId}`;
                const childContext = { ...childPayload, Id: parentId, Title: childTitle };

                void NotificationService.logNotification(
                  this.service.siteUrl,
                  childListKey,
                  childAction,
                  childContext,
                  this.state.currentUser,
                  config
                ).catch(e => console.error(`Child Notification Error (${childListKey})`, e));
              }
            }
          }
        }));
      }

      // --- STEP D: FINISH ---
      void Swal.fire({ icon: 'success', title: 'Success', text: 'Saved successfully!' });

      // RESET FORM STATE
      this.setState({
        mode: 'list',
        itemId: undefined,
        formData: {},
        childItems: {},
        attachmentsNew: [],
        attachmentsDelete: [],
        existingAttachments: []
      });
      await this.loadItems();

    } catch (error: any) {
      console.error("Submit Error:", error);
      let msg = error.message;
      if (error.data && error.data.responseBody && error.data.responseBody['odata.error']) {
        msg = error.data.responseBody['odata.error'].message.value;
      }
      void Swal.fire({ icon: 'error', title: 'Save Failed', html: `Error: ${msg}` });
    }
  }
  // --- BULK DELETE HANDLER ---
  private handleBulkDelete = async (): Promise<void> => {
    const { selectedItems } = this.state;

    // 1. Safety Checks
    if (!this.state.canDelete || this.props.overrideDelete) return;
    if (!selectedItems || selectedItems.length === 0) return;

    // 2. Confirmation
    const result = await Swal.fire({
      title: 'Are you sure?',
      text: `Delete ${selectedItems.length} item(s)?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
      confirmButtonText: 'Yes, delete'
    } as any);

    if (!result.isConfirmed && !(result as any).value) return;

    // 3. Processing
    void Swal.fire({
      title: 'Deleting...',
      html: 'Processing items...',
      allowOutsideClick: false,
      didOpen: () => { Swal.showLoading(); }
    });

    try {
      const list = this._sp.web.lists.getByTitle(this.props.selectedList);

      // Execute Deletes in Parallel
      const deletePromises = selectedItems.map(id => {
        const safeId = parseInt(String(id), 10);
        return list.items.getById(safeId).recycle()
          .catch(e => {
            console.warn(`Failed to recycle item ${safeId}`, e);
          });
      });

      await Promise.all(deletePromises);

      // 4. Success Message
      void Swal.fire({
        icon: 'success',
        title: 'Deleted!',
        text: `${selectedItems.length} items have been recycled.`,
        timer: 1500,
        showConfirmButton: false
      });

      // --- NOTIFICATION LOGIC ---
      if (this.props.enableNotification) {

        const config = this.getNotificationConfig();
        const listName = this.props.listPageTitle || this.props.selectedList || "List";

        // Notification Loop
        selectedItems.forEach(id => {
          // ---  Use Index instead of .find() ---
          const allIds = this.state.items.map(i => i.Id);
          const index = allIds.indexOf(id);
          const deletedItem = index > -1 ? this.state.items[index] : null;
          // ----------------------------------------

          const title = deletedItem ? deletedItem.Title : `Item ${id}`;

          void NotificationService.logNotification(
            this.service.siteUrl,
            listName,
            'Deleted',
            { Title: title, Id: id },
            this.state.currentUser,
            config
          );
        });

      }
      // 1. Update Global Cache (Prevents loadItems from restoring deleted data)
      this.allSourceItems = this.allSourceItems.filter(item => selectedItems.indexOf(item.Id) === -1);

      // 2. Update Current View
      const remainingItems = this.state.items.filter(item => selectedItems.indexOf(item.Id) === -1);

      // 3. Update State IMMEDIATELY
      this.setState({
        mode: 'list',
        itemId: undefined,
        items: remainingItems, // <--- Updates the grid
        selectedItems: []      // <--- Clears checkboxes
      });



    } catch (error: any) {
      let errMsg = error.message;
      if (error.data && error.data.responseBody && error.data.responseBody['odata.error']) {
        errMsg = error.data.responseBody['odata.error'].message.value;
      }
      void Swal.fire({ icon: 'error', title: 'Error', html: errMsg });
    }
  }
  // Handler: Add empty row for Grid Mode
  private addChildItemEmpty(listKey: string) {
    this.setState(prev => ({
      childItems: {
        ...prev.childItems,
        [listKey]: [...(prev.childItems[listKey] || []), {}] // Add empty object
      }
    }));
  }

  // Handler: Update Grid Input
  private updateChildItem(listKey: string, index: number, fieldName: string, value: any) {
    const items = [...(this.state.childItems[listKey] || [])];
    items[index] = { ...items[index], [fieldName]: value };

    this.setState(prev => ({
      childItems: { ...prev.childItems, [listKey]: items }
    }));
  }

  // Helper: Render Input for Grid
  private renderChildInput(listKey: string, index: number, field: ColumnDefinition, value: any) {
    // Simplified version of your renderFormField logic
    // For grid, usually just Text/Number/Dropdowns are best. 
    // Complex fields like PeoplePicker might be too big for a cell.

    return (
      <input
        value={value || ''}
        onChange={(e) => this.updateChildItem(listKey, index, field.EntityPropertyName, e.target.value)}
        className={styles.gridInput} // Add CSS for 100% width, no border
        style={{ width: '100%', border: '1px solid #eee', padding: 4 }}
      />
    );
  }
  private normaliseItemForForm(item: any): any {
    try {
      const normalised: any = {};
      for (let i = 0; i < this.state.fields.length; i++) {
        const f = this.state.fields[i];
        const name = f.EntityPropertyName;
        if (f.TypeAsString === 'Lookup' || f.TypeAsString === 'LookupMulti' || f.TypeAsString === 'User' || f.TypeAsString === 'UserMulti') {
          const idProp = f.InternalName + 'Id';
          if (f.AllowMultipleValues) {
            const mv = item[idProp] && item[idProp].results ? item[idProp].results : [];
            normalised[name] = mv;
          } else {
            normalised[name] = item[idProp];
          }
          continue;
        } else if (f.TypeAsString === 'MultiChoice' || (f.TypeAsString === 'Choice' && f.AllowMultipleValues)) {
          const mc = item[name];
          if (mc && mc.results) {
            normalised[name] = mc.results;
          } else if (typeof mc === 'string') {
            normalised[name] = mc.split(';#').filter((s: string) => s);
          } else {
            normalised[name] = [];
          }
        } else {
          normalised[name] = item[name];
        }
      }
      normalised['Created'] = item.Created;
      normalised['Modified'] = item.Modified;
      normalised['Author'] = item.Author;
      normalised['Editor'] = item.Editor;
      return normalised;
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - normaliseItemForForm',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return {}; // Return empty object to prevent downstream crashes
    }
  }

  private executeCustomScript(mode: string) {
    // ---------------------------------------------------------
    // 1. HANDLE CUSTOM JAVASCRIPT  
    // ---------------------------------------------------------
    let scriptContent = '';
    if (mode === 'list') scriptContent = this.props.listCustomScript ?? "";
    else if (mode === 'add') scriptContent = this.props.addCustomScript ?? "";
    else if (mode === 'edit') scriptContent = this.props.editCustomScript ?? "";
    else if (mode === 'view') scriptContent = this.props.viewCustomScript ?? "";

    if (scriptContent) {
      setTimeout(() => {
        try {
          const lines = scriptContent.split('\n');
          lines.forEach(line => {
            line = line.trim();
            if (!line) return;
            const isHttp = line.indexOf('http://') === 0;
            const isHttps = line.indexOf('https://') === 0;
            const isProtocolRelative = line.indexOf('//') === 0;
            const isJsFile = line.indexOf('.js') === (line.length - 3);

            if ((isHttp || isHttps || isProtocolRelative) && isJsFile) {
              const s = document.createElement('script');
              s.src = line;
              s.async = false;
              document.head.appendChild(s);
            } else {
              try {
                const FuncBuilder = (window as any).Function;
                const fn = new FuncBuilder(line);
                fn();
              } catch (error: any) {
                void LoggerService.log('PowerForm - Custom JS Error', 'Medium', 'N/A', error.message);
              }
            }
          });
        } catch (error: any) {
          void LoggerService.log('PowerForm - Custom JS Error', 'Medium', 'N/A', error.message);
        }
      }, 500);
    }

    // ---------------------------------------------------------
    // 2. HANDLE CUSTOM CSS (New Logic)
    // ---------------------------------------------------------
    let styleContent = '';
    // Fetch the correct style property based on the current mode
    if (mode === 'list') styleContent = this.props.listCustomStyle ?? "";
    else if (mode === 'add') styleContent = this.props.addCustomStyle ?? "";
    else if (mode === 'edit') styleContent = this.props.editCustomStyle ?? "";
    else if (mode === 'view') styleContent = this.props.viewCustomStyle ?? "";

    // A. CLEANUP: Remove previous styles to prevent bleeding between views
    const existingStyle = document.getElementById('PowerForm-custom-style-block');
    if (existingStyle) existingStyle.remove();

    const existingLink = document.getElementById('PowerForm-custom-css-link');
    if (existingLink) existingLink.remove();

    // B. INJECTION: Apply new styles if they exist
    if (styleContent) {
      try {
        const trimmed = styleContent.trim();
        // Check if it's a URL (starts with http/https or // and ends with .css)
        const isUrl = (trimmed.indexOf('http') === 0 || trimmed.indexOf('//') === 0) && trimmed.indexOf('.css') > -1;

        if (isUrl) {
          // Case 1: External CSS File
          const link = document.createElement('link');
          link.id = 'PowerForm-custom-css-link';
          link.rel = 'stylesheet';
          link.type = 'text/css';
          link.href = trimmed;
          document.head.appendChild(link);
        } else {
          // Case 2: Inline CSS Block
          const style = document.createElement('style');
          style.id = 'PowerForm-custom-style-block';
          style.type = 'text/css';
          style.appendChild(document.createTextNode(trimmed));
          document.head.appendChild(style);
        }
      } catch (error: any) {
        void LoggerService.log('PowerForm - Custom CSS Error', 'Medium', 'N/A', error.message);
      }
    }
  }
  private handleFilterChange = (field: string, value: string): void => {
    try {
      this.setState(
        prev => ({
          filters: { ...prev.filters, [field]: { operator: prev.filters[field]?.operator || 'contains', value: String(value) } },
          page: 1
        }),
        () => void this.loadItems()
      );
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleFilterChange',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private clearFilters = (): void => {
    try {
      this.setState({ filters: {}, searchText: '', page: 1 }, () => void this.loadItems());
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - clearFilters',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private handleSort = (field: string) => {
    try {
      const direction = this.state.sortField === field && this.state.sortDirection === 'asc' ? 'desc' : 'asc';
      this.setState({ sortField: field, sortDirection: direction, page: 1 }, () => void this.loadItems());
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleSort',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  private handleDelete = async (id: number) => {
    try {
      this.allSourceItems = [];

      // ---  Use Index instead of .find() ---
      // 1. Get the array of IDs
      const allIds = this.state.items.map(i => i.Id);
      // 2. Find the index of the specific ID
      const index = allIds.indexOf(id);
      // 3. Get the item (check if index is valid)
      const itemToDelete = index > -1 ? this.state.items[index] : null;

      // Fallback title
      const itemTitle = itemToDelete ? itemToDelete.Title : `Item ${id}`;

      const result = await Swal.fire({
        title: 'Are you sure?',
        text: "You won't be able to revert this!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#ef4444',
        cancelButtonColor: '#3b82f6',
        confirmButtonText: 'Yes, delete it!'
      });

      if (result.value) {
        try {
          await this._sp.web.lists.getByTitle(this.props.selectedList).items.getById(id).delete();
          void Swal.fire('Deleted!', 'Your file has been deleted.', 'success');

          // ==========================================
          // --- NOTIFICATION LOGIC (DELETE SINGLE) ---
          // ==========================================
          if (this.props.enableNotification) {

            const config = this.getNotificationConfig();
            const listName = this.props.listPageTitle || this.props.selectedList || "List";

            void NotificationService.logNotification(
              this.service.siteUrl,
              listName,
              'Deleted',
              { Title: itemTitle, Id: id }, // Uses captured title
              this.state.currentUser,
              config
            );

          }
          // ==========================================

          void this.loadItems();
        } catch (error: any) {
          void Swal.fire('Error!', 'Failed to delete item.', 'error');
          void LoggerService.log(
            'PowerForm - handleDelete',
            'High',
            id.toString(),
            error.message || JSON.stringify(error)
          );
        }
      }
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleDelete (System)',
        'Medium',
        id.toString(),
        error.message || JSON.stringify(error)
      );
    }
  }
  // REMOVE new attachment before upload
  private removeNewAttachment(idx: number): void {
    try {
      const copy = this.state.attachmentsNew.slice();
      copy.splice(idx, 1);
      this.setState({ attachmentsNew: copy });
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - removeNewAttachment',
        'Low',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  // DELETE existing attachment
  private deleteAttachment(id: string): void {
    try {
      this.setState(prev => ({
        attachmentsDelete: [...prev.attachmentsDelete, id],
        existingAttachments: prev.existingAttachments.filter(a => a.FileName !== id),
        attachments: prev.attachments.filter(a => a.id !== id)
      }));
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - deleteAttachment',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  // NEW: Export to Excel Handler
  private handleExport = () => {
    try {
      const { selectedItems, items, fields } = this.state;
      // Check if specific items are selected
      const hasSelection = selectedItems && selectedItems.length > 0;
      const itemsToExport = hasSelection
        ? items.filter(i => selectedItems.indexOf(i.Id) > -1)
        : items;
      if (itemsToExport.length === 0) {
        void Swal.fire('Info', 'No items to export.', 'info');
        return;
      }
      // Get visible columns
      const listCols = this.props.listVisibleFields && this.props.listVisibleFields.length > 0
        ? [...this.props.listVisibleFields]
        : (fields ? fields.map(f => f.InternalName) : []);
      const data = itemsToExport.map(item => {
        const row: any = {};
        listCols.forEach(fName => {
          const fDef = fields.find(f => f.InternalName === fName);
          const raw = item[fName];
          let val = '';
          if (!fDef) {
            // Fallback for unknown columns
            val = raw ? String(raw) : '';
          } else {
            //  Handle Note (Rich Text) - Strip HTML tags
            if (fDef.TypeAsString === 'Note') {
              val = raw ? String(raw).replace(/<[^>]+>/g, '') : '';
            }
            //  Handle Hyperlink (URL)
            else if (fDef.TypeAsString === 'URL') {
              if (raw && raw.Url) {
                // Export format: "Description (URL)" or just URL
                val = raw.Description ? `${raw.Description} (${raw.Url})` : raw.Url;
              }
            }
            //  Handle Multi-User / Multi-Lookup / Multi-Choice
            else if (['UserMulti', 'LookupMulti', 'MultiChoice'].indexOf(fDef.TypeAsString) > -1) {
              // Check if it's inside a 'results' object (standard OData) or direct array
              let arr = [];
              if (Array.isArray(raw)) {
                arr = raw;
              } else if (raw && raw.results && Array.isArray(raw.results)) {
                arr = raw.results;
              }
              // Map to comma-separated string
              if (arr.length > 0) {
                val = arr.map((r: any) => r.Title || r).join(', ');
              }
            }
            //  Handle Single User / Lookup
            else if (fDef.TypeAsString === 'User' || fDef.TypeAsString === 'Lookup') {
              if (raw && raw.Title) val = raw.Title;
              else if (raw) val = String(raw); // Fallback
            }
            //  Default Handler (Text, Number, Date, etc.)
            else {
              if (raw && typeof raw === 'object' && raw.Title) {
                val = raw.Title;
              } else {
                val = raw !== null && raw !== undefined ? String(raw) : '';
              }
            }
          }
          row[fName] = val;
        });
        return row;
      });
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Items');
      XLSX.writeFile(wb, `${this.props.selectedList}_export_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleExport',
        'High',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }

  // --- HANDLE CHECKBOX SELECTION (ES5 Compatible) ---
  private handleCheckboxChange = (id: number, checked: boolean): void => {
    this.setState((prevState) => {
      const selected = new Set(prevState.selectedItems);
      if (checked) {
        selected.add(id);
      } else {
        selected.delete(id);
      }

      //  Use forEach instead of Array.from() for ES5 compatibility
      const newItems: number[] = [];
      selected.forEach((item) => {
        newItems.push(item);
      });

      return { selectedItems: newItems };
    });
  }
  // NEW: Toggle Select All
  private handleSelectAll = (checked: boolean) => {
    try {
      if (checked) {
        // Select all IDs currently visible in the table
        const allIds = this.state.items.map(i => i.Id);
        this.setState({ selectedItems: allIds });
      } else {
        this.setState({ selectedItems: [] });
      }
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - handleSelectAll',
        'Low',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
    }
  }
  // NEW: Get CSS class for status tag
  private getTagClass(status: string): string {
    try {
      if (!status) return '';
      const s = String(status).toLowerCase();
      if (s === 'active' || s === 'approved' || s === 'completed') return styles.success;
      if (s === 'draft' || s === 'pending') return '';
      if (s === 'rejected' || s === 'low stock' || s === 'issue') return styles.warn;
      if (s === 'deleted' || s === 'error') return styles.danger;
      return '';
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - getTagClass',
        'Low',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return ''; // Return empty string on error so UI doesn't break
    }
  }
  // NEW: Helper for Standard Synchronous Validations (Required, Min/Max)
  private getStandardValidation(field: ColumnDefinition, value: any): string {
    try {
      //  Required Check
      if (field.Required) {
        if (value === null || value === undefined || value === '') {
          return `${field.Title} is required.`;
        }
        if (Array.isArray(value) && value.length === 0) {
          return `${field.Title} is required.`;
        }
      }
      //  Number / Currency Range Check
      if (field.TypeAsString === 'Number' || field.TypeAsString === 'Currency') {
        const num = parseFloat(value);
        // Only validate range if it is a valid number
        if (!isNaN(num)) {
          // Check Minimum
          if (field.MinimumValue !== undefined && field.MinimumValue !== null) {
            if (num < field.MinimumValue) {
              return `Value must be at least ${field.MinimumValue}.`;
            }
          }
          // Check Maximum
          if (field.MaximumValue !== undefined && field.MaximumValue !== null) {
            if (num > field.MaximumValue) {
              return `Value must be at most ${field.MaximumValue}.`;
            }
          }
        }
      }
      return '';
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - getStandardValidation',
        'Medium',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return ''; // Fail safe: Return no error if validation crashes
    }
  }
  // NEW: Render Pagination Controls
  private renderPagination(): React.ReactElement<any> {
    try {
      //  Get State (use totalItems from state for server-side paging accuracy)
      const { page, pageSize, totalItems: stateTotalItems } = this.state;
      // Fallback to items.length if totalItems not available (client-side compatibility)
      const totalItems = stateTotalItems > 0 ? stateTotalItems : this.state.items.length;
      //  Calculate Ranges
      const totalPages = Math.ceil(totalItems / pageSize) || 1;
      const start = totalItems === 0 ? 0 : (page - 1) * pageSize + 1;
      const end = Math.min(page * pageSize, totalItems);
      //  Generate Visible Page Numbers (sliding window of 5)
      const maxVisible = 5;
      let startP = Math.max(1, page - Math.floor(maxVisible / 2));
      let endP = Math.min(totalPages, startP + maxVisible - 1);
      if (endP - startP + 1 < maxVisible) {
        startP = Math.max(1, endP - maxVisible + 1);
      }
      const pagesToShow: number[] = [];
      for (let i = startP; i <= endP; i++) {
        pagesToShow.push(i);
      }
      //  Ellipsis Logic
      const showLeftEllipsis = startP > 1;
      const showRightEllipsis = endP < totalPages;
      //  Return JSX
      return (
        <div className={styles.paginationFooter}>
          <div className={styles.paginationInfo}>
            Showing {start} to {end} of {totalItems} items
          </div>
          <div className={styles.paginationControls}>
            {/* First Page */}
            <button
              className={styles.btn}
              disabled={page === 1}
              onClick={() => this.setState({ page: 1 })}
            >
              |&lt;
            </button>
            {/* Previous Page */}
            <button
              className={styles.btn}
              disabled={page === 1}
              onClick={() => this.setState({ page: page - 1 })}
            >
              &lt;&lt;
            </button>
            {/* Left Ellipsis */}
            {showLeftEllipsis && <span style={{ padding: '0 8px', color: '#999' }}>...</span>}
            {/* Page Numbers */}
            {pagesToShow.map(p => (
              <button
                key={p}
                className={`${styles.btn} ${p === page ? styles.active : ''}`}
                onClick={() => this.setState({ page: p })}
              >
                {p}
              </button>
            ))}
            {/* Right Ellipsis */}
            {showRightEllipsis && <span style={{ padding: '0 8px', color: '#999' }}>...</span>}
            {/* Next Page */}
            <button
              className={styles.btn}
              disabled={page === totalPages}
              onClick={() => this.setState({ page: page + 1 })}
            >
              &gt;&gt;
            </button>
            {/* Last Page */}
            <button
              className={styles.btn}
              disabled={page === totalPages}
              onClick={() => this.setState({ page: totalPages })}
            >
              &gt;|
            </button>
          </div>
          <div className={styles.perPage}>
            <label>Rows per page:</label>
            <select
              value={pageSize}
              onChange={(e) => this.setState({ pageSize: parseInt(e.target.value, 10), page: 1 })}
            >
              <option value="5">5</option>
              <option value="10">10</option>
              <option value="25">25</option>
              <option value="50">50</option>
              <option value="100">100</option>
            </select>
          </div>
        </div>
      );
    } catch (error: any) {
      void LoggerService.log(
        'PowerForm - renderPagination',
        'Low',
        this.state.itemId ? this.state.itemId.toString() : 'N/A',
        error.message || JSON.stringify(error)
      );
      return <React.Fragment />; // Return nothing if pagination crashes
    }
  }
  // NEW: Render Form Field based on type and mode
  private renderFormField(field: ColumnDefinition, isView: boolean, isReadOnly: boolean) {
    try {
      const key = field.EntityPropertyName;
      const value = this.state.formData[key];
      const isDateOnly = field.DisplayFormat === 0;
      // 1. DETERMINE INDIVIDUAL LOCK FLAGS FIRST
      let forceReadOnly = false;
      if (this.props.fieldPermissionConfig && this.props.fieldPermissionConfig[key]) {
        const allowedGroups = this.props.fieldPermissionConfig[key];
        if (allowedGroups.length > 0) {
          const userGroups = this.state.currentUserGroups || [];
          const hasAccess = allowedGroups.some(g => userGroups.indexOf(g) > -1);
          if (!hasAccess) forceReadOnly = true;
        }
      }
      const isUrlLocked = this.state.urlReadOnlyFields && this.state.urlReadOnlyFields.indexOf(key) > -1;
      // 2. CALCULATE COMPOSITE RESTRICTION FLAG 
      const isRestricted = forceReadOnly || isReadOnly || isUrlLocked;
      const restrictionMessage = "Restricted by administrator. Please contact him for more info.";
      const error = this.state.formErrors[key];
      const effectiveReadOnly = isView || isRestricted;
      const repConfig = this.props.repeaterConfig && (this.props.repeaterConfig[field.InternalName] || this.props.repeaterConfig[field.EntityPropertyName]);

      // 3. RENDER HELPERS
      const renderRestrictedIcon = () => (
        isRestricted ? (
          <span title={restrictionMessage} style={{ marginLeft: '8px', cursor: 'help', color: '#d13438' }}>
            <i className="ms-Icon ms-Icon--Lock" aria-hidden="true" style={{ fontSize: '12px' }}></i>
          </span>
        ) : null
      );

      // Style for grayed out controls 
      const restrictedStyle: React.CSSProperties = isRestricted ? {
        backgroundColor: '#f3f2f1',
        cursor: 'not-allowed',
        color: '#a19f9d',
        border: '1px solid #c8c6c4'
      } : {};
      if (isView) {
        if (repConfig && field.TypeAsString === 'Note') {
          return (
            <div key={key} className={styles.field}>
              <label>{field.Title}</label>
              {/* Render Repeater in Read-Only Mode */}
              <RepeaterInput
                columns={repConfig}
                value={value || '[]'}
                mode="view"
                onChange={() => { }} // No-op for view
              />
            </div>
          );
        }
        if (field.TypeAsString === 'Note') return <div key={key} className={styles.field}><label>{field.Title}</label><div className={styles.viewField} dangerouslySetInnerHTML={{ __html: value || '-' }}></div></div>;
        if (field.TypeAsString === 'URL') { const u = value || { Url: '', Description: '' }; return <div key={key} className={styles.field}><label>{field.Title}</label><div className={styles.viewField}><a href={u.Url} target="_blank">{u.Description || u.Url}</a></div></div>; }
        if (field.TypeAsString === 'Attachments') {
          const attEls = (this.state.attachments && this.state.attachments.length > 0) ? (
            this.state.attachments.map((a: any) => (
              <div key={a.id} style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '4px' }}>
                <Icons.Clip />
                <a href={a.url} target="_blank" rel="noopener noreferrer" style={{ textDecoration: 'none', color: '#3b82f6' }}>{a.name}</a>
              </div>
            ))
          ) : <div>None</div>;
          return <div key={key} className={styles.field}><label>{field.Title}</label><div className={styles.viewField}>{attEls}</div></div>;
        }
        let displayValue: any = value;
        if (field.TypeAsString === 'MultiChoice' && Array.isArray(value)) displayValue = value.join(', ');
        if (field.TypeAsString === 'Boolean') displayValue = value ? 'Yes' : 'No';
        if (field.TypeAsString === 'DateTime') displayValue = this.formatDate_form(value, isDateOnly);
        if (field.TypeAsString.indexOf('User') > -1 || field.TypeAsString.indexOf('Lookup') > -1) {
          if (this.state.lookupOptions[key]) {
            if (Array.isArray(value)) displayValue = value.map(v => { const o = this.state.lookupOptions[key].find((op: any) => op.key === v); return o ? o.text : v; }).join(', ');
            else { const found = this.state.lookupOptions[key].find((op: any) => op.key === value); displayValue = found ? found.text : value; }
          }
          if (field.TypeAsString.indexOf('User') > -1) {
            const userIds = Array.isArray(value) ? value : (value ? [value] : []);
            displayValue = (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                {userIds.map(uid => {
                  const persona = this.state.peopleOptions[uid];
                  if (!persona) return <div key={uid}>{uid}</div>;
                  const userName = persona.text || persona.primaryText;
                  const userEmail = persona.secondaryText;
                  const photoUrl = `${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${userEmail}`;
                  return (
                    <div key={uid} style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <img src={photoUrl} alt={userName} style={{ width: '32px', height: '32px', borderRadius: '50%', objectFit: 'cover' }} onError={(e: any) => { e.target.src = "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-default.png"; }} />
                      <div style={{ display: 'flex', flexDirection: 'column' }}>
                        <span style={{ fontWeight: 600, color: '#333' }}>{userName}</span>
                        {userEmail && <span style={{ fontSize: '11px', color: '#666' }}>{userEmail}</span>}
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          }
        }
        return <div key={key} className={styles.field}><label>{field.Title}</label><div className={styles.viewField}>{typeof displayValue === 'object' && !React.isValidElement(displayValue) ? JSON.stringify(displayValue) : (displayValue || '-')}</div></div>;
      }
      switch (field.TypeAsString) {
        case 'Note': {
          if (repConfig) {
            return (
              <div key={key} className={styles.field} title={isRestricted ? restrictionMessage : ""}>
                <label>
                  {field.Title} {field.Required && !effectiveReadOnly && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}
                </label>
                <RepeaterInput columns={repConfig} value={value || '[]'} mode={effectiveReadOnly ? 'view' : 'edit'} onChange={(newJson) => this.handleChange(key, newJson)} />
                {error && !effectiveReadOnly && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
              </div>
            );
          }
          const isRichText = (field as any).RichText === true;
          return (
            <div key={key} className={styles.field} title={isRestricted ? restrictionMessage : ""} style={{ display: 'flex', flexDirection: 'column', gridRow: 'span 2', height: '100%' }}>
              <label>{field.Title} {field.Required && !effectiveReadOnly && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
              <div style={{ flexGrow: 1, ...restrictedStyle }}>
                {isRichText ? (
                  <RichTextEditor value={value || ''} readOnly={effectiveReadOnly} onChange={(val: string) => this.handleChange(key, val)} />
                ) : (
                  <textarea rows={6} disabled={effectiveReadOnly} value={value || ''} onChange={(e) => this.handleChange(key, e.target.value)} onBlur={() => void this.validateField(field, value)} style={{ resize: 'vertical', width: '100%', minHeight: '100px', ...restrictedStyle }} />
                )}
              </div>
              {error && !effectiveReadOnly && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
            </div>
          );
        }
        case 'URL': {
          const urlVal = value || { Url: '', Description: '' };
          return (
            <div key={key} className={styles.field} title={isRestricted ? restrictionMessage : ""} style={{ border: '1px solid #eee', padding: 10, borderRadius: 5, background: effectiveReadOnly ? '#f9fafb' : '#fff', ...restrictedStyle }}>
              <label>{field.Title} {field.Required && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
              <div style={{ display: 'grid', gap: 5 }}>
                <input disabled={effectiveReadOnly} placeholder="URL" value={urlVal.Url} onChange={(e) => this.handleChange(key, { Url: e.target.value, Description: urlVal.Description })} style={restrictedStyle} />
                <input disabled={effectiveReadOnly} placeholder="Description" value={urlVal.Description} onChange={(e) => this.handleChange(key, { Url: urlVal.Url, Description: e.target.value })} style={restrictedStyle} />
              </div>
            </div>
          );
        }
        case 'Choice':
        case 'MultiChoice': {
          const isMultiChoice = field.TypeAsString === 'MultiChoice';
          const allChoices = field.Choices || [];
          const isOpen = this.state.activePickerKey === key;
          let displayValue = '';
          let currentVals: string[] = [];
          if (isMultiChoice) { currentVals = Array.isArray(value) ? value : []; displayValue = currentVals.join(', '); }
          else if (value) { currentVals = [String(value)]; displayValue = String(value); }
          const searchText = ((this.state.pickerSearch && this.state.pickerSearch[key]) || '').toLowerCase();
          const filteredChoices = allChoices.filter(c => c && c.toLowerCase().indexOf(searchText) > -1);
          return (
            <div key={key} className={styles.field} style={{ position: 'relative' }} data-picker-wrapper={key} title={isRestricted ? restrictionMessage : ""}>
              <label>{field.Title} {field.Required && !effectiveReadOnly && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
              <div onClick={() => { if (!effectiveReadOnly) this.setState({ activePickerKey: isOpen ? null : key, pickerSearch: { ...this.state.pickerSearch, [key]: '' } }); }} style={{ border: '1px solid #e5e7eb', borderRadius: '6px', padding: '8px 10px', background: isRestricted ? '#f3f2f1' : '#fff', cursor: isRestricted ? 'not-allowed' : 'pointer', minHeight: '36px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', ...restrictedStyle }}>
                <span style={{ color: displayValue ? '#374151' : '#9ca3af', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{displayValue || (isMultiChoice ? "Select options..." : "Select an option...")}</span>
                {isRestricted ? <i className="ms-Icon ms-Icon--Lock" style={{ fontSize: '12px' }}></i> : <span style={{ fontSize: '10px', color: '#6b7280' }}>▼</span>}
              </div>
              {isOpen && !effectiveReadOnly && (
                <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 1000, background: '#fff', border: '1px solid #e5e7eb', borderRadius: '6px', marginTop: '4px', boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)', padding: '8px' }}>
                  <input type="text" placeholder="Search..." autoFocus value={searchText} onChange={(e) => this.setState({ pickerSearch: { ...this.state.pickerSearch, [key]: e.target.value } })} style={{ width: '100%', padding: '6px 8px', marginBottom: '6px', border: '1px solid #d1d5db', borderRadius: '4px', fontSize: '13px', outline: 'none' }} onClick={(e) => e.stopPropagation()} />
                  <div style={{ maxHeight: '180px', overflowY: 'auto' }}>
                    {filteredChoices.map((choice, idx) => {
                      const isSelected = currentVals.indexOf(choice) > -1;
                      return (
                        <div key={idx} onClick={(e) => { e.stopPropagation(); let newVal: any = isMultiChoice ? (isSelected ? currentVals.filter(v => v !== choice) : [...currentVals, choice]) : choice; this.handleChange(key, newVal); if (!isMultiChoice) this.setState({ activePickerKey: null }); }} style={{ padding: '6px 8px', cursor: 'pointer', fontSize: '13px', borderRadius: '4px', background: isSelected ? '#eff6ff' : 'transparent', color: isSelected ? '#1e3a8a' : '#374151', display: 'flex', alignItems: 'center' }}>
                          <div style={{ width: '20px', marginRight: '4px' }}>{isSelected && <span style={{ color: '#2563eb', fontWeight: 'bold' }}>✓</span>}</div>{choice}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
              {error && !effectiveReadOnly && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
            </div>
          );
        }
        case 'Lookup':
        case 'LookupMulti': {
          const isLookupMulti = field.TypeAsString === 'LookupMulti';
          const rawOpts = (this.state.lookupOptions[key] || []);
          const isOpenLookup = this.state.activePickerKey === key;
          let displayLookupValue = '';
          let currentIds: number[] = [];
          if (isLookupMulti) { currentIds = Array.isArray(value) ? value : []; displayLookupValue = currentIds.map(id => { const m = rawOpts.find((o: any) => o.key === id); return m ? m.text : id; }).join(', '); }
          else if (value) { currentIds = [Number(value)]; const m = rawOpts.find((o: any) => o.key === value); displayLookupValue = m ? m.text : value; }
          const searchLookupText = ((this.state.pickerSearch && this.state.pickerSearch[key]) || '').toLowerCase();
          const filteredLookupOpts = rawOpts.filter((o: any) => o && o.text && o.text.toLowerCase().indexOf(searchLookupText) > -1);
          return (
            <div key={key} className={styles.field} style={{ position: 'relative' }} data-picker-wrapper={key} title={isRestricted ? restrictionMessage : ""}>
              <label>{field.Title} {field.Required && !effectiveReadOnly && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
              <div onClick={() => { if (!effectiveReadOnly) this.setState({ activePickerKey: isOpenLookup ? null : key, pickerSearch: { ...this.state.pickerSearch, [key]: '' } }); }} style={{ border: '1px solid #e5e7eb', borderRadius: '6px', padding: '8px 10px', background: isRestricted ? '#f3f2f1' : '#fff', cursor: isRestricted ? 'not-allowed' : 'pointer', minHeight: '36px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', ...restrictedStyle }}>
                <span style={{ color: displayLookupValue ? '#374151' : '#9ca3af', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{displayLookupValue || (isLookupMulti ? "Select items..." : "Select an item...")}</span>
                {isRestricted ? <i className="ms-Icon ms-Icon--Lock" style={{ fontSize: '12px' }}></i> : <span style={{ fontSize: '10px', color: '#6b7280' }}>▼</span>}
              </div>
              {isOpenLookup && !effectiveReadOnly && (
                <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 1000, background: '#fff', border: '1px solid #e5e7eb', borderRadius: '6px', marginTop: '4px', boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)', padding: '8px' }}>
                  <input type="text" placeholder="Search..." autoFocus value={searchLookupText} onChange={(e) => this.setState({ pickerSearch: { ...this.state.pickerSearch, [key]: e.target.value } })} style={{ width: '100%', padding: '6px 8px', marginBottom: '6px', border: '1px solid #d1d5db', borderRadius: '4px', fontSize: '13px', outline: 'none' }} onClick={(e) => e.stopPropagation()} />
                  <div style={{ maxHeight: '180px', overflowY: 'auto' }}>
                    {filteredLookupOpts.map((opt: any) => {
                      const isSelected = currentIds.indexOf(Number(opt.key)) > -1;
                      return (
                        <div key={opt.key} onClick={(e) => { e.stopPropagation(); let newVal: any = isLookupMulti ? (isSelected ? currentIds.filter(id => id !== Number(opt.key)) : [...currentIds, Number(opt.key)]) : Number(opt.key); this.handleChange(key, newVal, opt.itemData); if (!isLookupMulti) this.setState({ activePickerKey: null }); }} style={{ padding: '6px 8px', cursor: 'pointer', fontSize: '13px', borderRadius: '4px', background: isSelected ? '#eff6ff' : 'transparent', color: isSelected ? '#1e3a8a' : '#374151', display: 'flex', alignItems: 'center' }}>
                          <div style={{ width: '20px', marginRight: '4px' }}>{isSelected && <span style={{ color: '#2563eb', fontWeight: 'bold' }}>✓</span>}</div>{opt.text}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
              {error && !effectiveReadOnly && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
            </div>
          );
        }
        case 'Boolean': {
          const isOpenBool = this.state.activePickerKey === key;
          const displayBoolValue = value === true ? 'Yes' : (value === false ? 'No' : '');
          const searchBoolVal = ((this.state.pickerSearch && this.state.pickerSearch[key]) || '').toLowerCase();
          const boolOptions = ['Yes', 'No'].filter(opt => opt.toLowerCase().indexOf(searchBoolVal) > -1);
          return (
            <div key={key} className={styles.field} style={{ position: 'relative' }} data-picker-wrapper={key} title={isRestricted ? restrictionMessage : ""}>
              <label>{field.Title} {field.Required && !effectiveReadOnly && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
              <div onClick={() => { if (!effectiveReadOnly) this.setState({ activePickerKey: isOpenBool ? null : key, pickerSearch: { ...this.state.pickerSearch, [key]: '' } }); }} style={{ border: '1px solid #e5e7eb', borderRadius: '6px', padding: '8px 10px', background: isRestricted ? '#f3f2f1' : '#fff', cursor: isRestricted ? 'not-allowed' : 'pointer', minHeight: '36px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', ...restrictedStyle }}>
                <span style={{ color: displayBoolValue ? '#374151' : '#9ca3af' }}>{displayBoolValue || "Select Yes/No..."}</span>
                {isRestricted ? <i className="ms-Icon ms-Icon--Lock" style={{ fontSize: '12px' }}></i> : <span style={{ fontSize: '10px', color: '#6b7280' }}>▼</span>}
              </div>
              {isOpenBool && !effectiveReadOnly && (
                <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 1000, background: '#fff', border: '1px solid #e5e7eb', borderRadius: '6px', marginTop: '4px', boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)', padding: '8px' }}>
                  <input type="text" placeholder="Search..." autoFocus value={searchBoolVal} onChange={(e) => this.setState({ pickerSearch: { ...this.state.pickerSearch, [key]: e.target.value } })} style={{ width: '100%', padding: '6px 8px', marginBottom: '6px', border: '1px solid #d1d5db', borderRadius: '4px', fontSize: '13px', outline: 'none' }} onClick={(e) => e.stopPropagation()} />
                  <div style={{ maxHeight: '180px', overflowY: 'auto' }}>
                    {boolOptions.map(opt => (
                      <div key={opt} onClick={(e) => { e.stopPropagation(); this.handleChange(key, opt === 'Yes'); this.setState({ activePickerKey: null }); }} style={{ padding: '6px 8px', cursor: 'pointer', fontSize: '13px', borderRadius: '4px', background: (opt === 'Yes' && value === true) || (opt === 'No' && value === false) ? '#eff6ff' : 'transparent', color: '#374151', display: 'flex', alignItems: 'center' }}>
                        <div style={{ width: '20px', marginRight: '4px' }}>{(opt === 'Yes' && value === true) || (opt === 'No' && value === false) ? <span style={{ color: '#2563eb', fontWeight: 'bold' }}>✓</span> : null}</div>{opt}
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {error && !effectiveReadOnly && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
            </div>
          );
        }
        case 'DateTime': {
          return (
            <div key={key} className={styles.field} title={isRestricted ? restrictionMessage : ""}>
              <label>{field.Title} {field.Required && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
              <input type={field.DisplayFormat === 0 ? "date" : "datetime-local"} disabled={effectiveReadOnly} value={this.formatDateTimeForInput(value, field.DisplayFormat === 0)} onChange={(e) => this.handleChange(key, e.target.value)} style={restrictedStyle} />
              {error && !effectiveReadOnly && <span style={{ color: 'red', fontSize: 12, display: 'block', marginTop: 4 }}>{error}</span>}
            </div>
          );
        }
       
        case 'User':
        case 'UserMulti': {
          const isUserMulti = field.TypeAsString === 'UserMulti' || field.AllowMultipleValues;
          const selectedIds = value ? (Array.isArray(value) ? value : [value]) : [];
          const selectedPersonas = selectedIds.map((id: number) => this.state.peopleOptions[id]).filter((p: any) => !!p);
          
          return (
            <div key={key} className={styles.field} title={isRestricted ? restrictionMessage : ""}>
              <label>{field.Title} {field.Required && !effectiveReadOnly && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
              <div style={{ pointerEvents: effectiveReadOnly ? 'none' : 'auto', opacity: effectiveReadOnly ? 0.6 : 1, ...restrictedStyle, padding: '4px', borderRadius: '4px' }}>
                <NormalPeoplePicker 
                  onResolveSuggestions={this._resolvePeopleSuggestions} 
                  pickerSuggestionsProps={{ suggestionsHeaderText: 'People & Groups', noResultsFoundText: 'No results found', loadingText: 'Loading' }} 
                  itemLimit={isUserMulti ? undefined : 1} 
                  selectedItems={selectedPersonas} 
                  onChange={(items?: IPersonaProps[]) => { 
                    const currentItems = items || []; 
                    const newIds = currentItems.map(i => parseInt((i.key as string) ?? "0")); 
                    const newOptions = { ...this.state.peopleOptions }; 
                    currentItems.forEach(i => { 
                      const userId = parseInt((i.key as string) ?? "0"); 
                      if (userId) newOptions[userId] = i as any; 
                    }); 
                    this.setState({ peopleOptions: newOptions }); 
                    this.handleChange(key, isUserMulti ? newIds : (newIds[0] || null)); 
                  }} 
                />
              </div>
              {error && !effectiveReadOnly && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
            </div>
          );
        }

        case 'Attachments': {
          return (
            <div key={key} className={styles.field} title={isRestricted ? restrictionMessage : ""}>
              <label>{field.Title} {renderRestrictedIcon()}</label>
              {this.state.existingAttachments && (
                <ul style={{ listStyle: 'none', padding: 0, margin: '0 0 10px 0' }}>
                  {this.state.existingAttachments.map((file: any) => (
                    <li key={file.FileName} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4, fontSize: 13 }}>
                      <Icons.Clip /><a href={file.ServerRelativeUrl} target="_blank" style={{ color: '#0078d4', textDecoration: 'none' }}>{file.FileName}</a>
                      {!effectiveReadOnly && <button type="button" onClick={() => this.deleteAttachment(file.FileName)} style={{ background: 'transparent', border: 'none', color: 'red', cursor: 'pointer' }}>✕</button>}
                    </li>
                  ))}
                </ul>
              )}
              {!effectiveReadOnly && <input type="file" multiple onChange={this.handleFileChange} style={{ display: 'block', marginTop: 5, ...restrictedStyle }} />}
              {error && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
            </div>
          );
        }
        default: {
          let acConfig = this.props.autocompleteConfig && (this.props.autocompleteConfig[field.InternalName] || this.props.autocompleteConfig[field.EntityPropertyName]);
          if (acConfig && acConfig.sourceList && acConfig.sourceField && !effectiveReadOnly) {
            const selectedTags: ITag[] = value ? [{ key: String(value), name: String(value) }] : [];
            return (
              <div key={key} className={styles.field} title={isRestricted ? restrictionMessage : ""}>
                <label>{field.Title} {field.Required && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
                <div style={{ opacity: effectiveReadOnly ? 0.7 : 1, pointerEvents: effectiveReadOnly ? 'none' : 'auto', ...restrictedStyle }}>
                  <TagPicker onResolveSuggestions={(filter) => this.onAutocompleteSearch(filter, field.InternalName).then(opts => opts.map((o: any) => ({ key: String(o.key), name: o.text, itemData: o.data } as any)))} onRenderSuggestionsItem={(props: any) => { let displayText = props.name; if (acConfig.additionalFields && acConfig.additionalFields.length > 0 && props.itemData) { const extras = acConfig.additionalFields.map((f: string) => { let v = props.itemData[f] || props.itemData[this.toODataName(f)]; return (typeof v === 'object' && v ? v.Title : v) || ''; }).filter((v: any) => v !== '').join(' | '); if (extras) displayText = `${displayText} (${extras})`; } return <div style={{ padding: '6px 12px' }}>{displayText}</div>; }} itemLimit={1} selectedItems={selectedTags} onChange={(items?: ITag[]) => this.handleChange(key, items && items.length > 0 ? items[0].name : '', items && items.length > 0 ? (items[0] as any).itemData : null)} inputProps={{ placeholder: 'Type to search...' }} disabled={effectiveReadOnly} />
                </div>
                {error && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
              </div>
            );
          }
          return (
            <div key={key} className={styles.field} title={isRestricted ? restrictionMessage : ""}>
              <label>{field.Title} {field.Required && !isRestricted && <span style={{ color: 'red' }}>*</span>} {renderRestrictedIcon()}</label>
              <input type={field.TypeAsString === 'Number' || field.TypeAsString === 'Currency' ? 'number' : 'text'} disabled={effectiveReadOnly} value={value || ''} onChange={(e) => this.handleChange(key, e.target.value)} onBlur={() => void this.validateField(field, value)} style={restrictedStyle} />
              {error && !effectiveReadOnly && <span style={{ color: 'red', fontSize: 12 }}>{error}</span>}
            </div>
          );
        }
      }
    } catch (error: any) {
      void LoggerService.log('PowerForm - renderFormField', 'High', this.state.itemId ? this.state.itemId.toString() : 'N/A', `Field: ${field.InternalName} - ${error.message}`);
      return <div key={field.EntityPropertyName} className={styles.field}><label>{field.Title}</label><div style={{ color: 'red', padding: '10px', border: '1px solid red', borderRadius: '4px' }}>Error rendering this field.</div></div>;
    }
  }
  // MAIN RENDER METHOD
  public render(): React.ReactElement<IPowerFormProps> {
    const { mode, items, filters, searchText, fields, page, pageSize } = this.state;
    // This variable 'currentItems' now holds ONLY the 10 items for the current page
    let currentItems: any[] = [];
    if (this.props.isLargeList) {
      // LARGE LIST: Use items directly because SharePoint already paginated the data
      currentItems = items;
    } else {
      // UP TO 5K: Use existing client-side slicing logic
      const currPage = page || 1;
      const size = pageSize || 10;
      const indexOfLastItem = currPage * size;
      const indexOfFirstItem = indexOfLastItem - size;
      currentItems = items.slice(indexOfFirstItem, indexOfLastItem);
    }
    // -------------------------
    const isList = mode === 'list';
    const customActions = mode === 'edit'
      ? (this.props.editCustomActions || [])
      : mode === 'view'
        ? (this.props.viewCustomActions || [])
        : [];
    // ... (Your existing field logic logic stays here) ...
    let visibleFieldKeys: string[] = [];
    let fieldOrder: { [key: string]: number } = {};
    let readOnlyKeys: string[] = [];
    if (mode === 'add') {
      visibleFieldKeys = this.props.addVisibleFields || [];
      fieldOrder = this.props.addFieldOrder || {};
      readOnlyKeys = this.props.addReadOnlyFields || [];
    } else if (mode === 'edit') {
      visibleFieldKeys = this.props.editVisibleFields || [];
      fieldOrder = this.props.editFieldOrder || {};
      readOnlyKeys = this.props.editReadOnlyFields || [];
    } else if (mode === 'view') {
      visibleFieldKeys = this.props.viewVisibleFields || [];
      fieldOrder = this.props.viewFieldOrder || {};
    }
    const formFields = fields.filter(f => {
      // 1. Check if the field is in the visible keys array
      const isVisibleKey = visibleFieldKeys.indexOf(f.InternalName) > -1;

      // 2. Exclude standard system/audit fields by EntityPropertyName
      const isSystemField = ['Author', 'Editor', 'Created', 'Modified', 'ID', 'Guid', 'ComplianceAssetId', 'ContentType']
        .indexOf(f.EntityPropertyName) === -1;

      // 3. Apply SharePoint-specific visibility logic
      // Cast 'f' to 'any' to avoid TS7053 errors for 'Hidden' and 'ReadOnlyField'
      const isNotHidden = !(f as any)['Hidden'];

      // Allow the field if it's not Read-Only, OR if it's the 'Attachments' field (special exception)
      const isEditableOrAttachments = !(f as any)['ReadOnlyField'] || f.InternalName === 'Attachments';

      return isVisibleKey && isSystemField && isNotHidden && isEditableOrAttachments;
    });
    if (fieldOrder) {
      formFields.sort((a, b) => (fieldOrder[a.InternalName] || 999) - (fieldOrder[b.InternalName] || 999));
    }
    //  Determine Base Columns (From View OR Props)
    let listCols: string[] = [];
    if (this.state.activeViewFields && this.state.activeViewFields.length > 0) {
      // CASE A: Custom View is Active -> Use its columns
      listCols = [...this.state.activeViewFields];
    } else {
      // CASE B: Default View -> Use Property Pane Settings
      listCols = this.props.listVisibleFields && this.props.listVisibleFields.length > 0
        ? [...this.props.listVisibleFields]
        : (this.state.fields ? this.state.fields.filter(f => f.InternalName !== 'ContentType').map(f => f.InternalName) : []);
      // Apply Sort Order only for Default View (Views have their own fixed order)
      const listOrder = this.props.listFieldOrderMap;
      if (listOrder) {
        listCols.sort((a, b) => (listOrder[a] || 999) - (listOrder[b] || 999));
      }
    }
    const listOrder = this.props.listFieldOrderMap;
    if (listOrder) {
      listCols.sort((a, b) => (listOrder[a] || 999) - (listOrder[b] || 999));
    }
    const isSaveDisabled = this.isFormInvalid(formFields);
    const filterConfig = this.props.listFilterMap || {};
    const hasCustomFilters = Object.keys(filterConfig).length > 0;
    const renderUser = (user: any) => {
      if (!user) return '';
      // 1. EXTRACT NAME: Check PascalCase (Standard) and lowercase (Large List)
      const userName = user.Title || user.title ||
        user.lookupValue || user.value ||
        user.Author || user.Editor ||
        user.Name || 'Unknown';
      // 2. EXTRACT EMAIL: Check PascalCase and lowercase versions
      let userEmail = user.EMail || user.Email || user.email || '';
      // FALLBACK: Parse from claims if needed
      if (!userEmail && user.Name && user.Name.indexOf('|') > -1) {
        userEmail = user.Name.split('|').pop();
      }
      // 3. GENERATE PHOTO URL: Ensure the identifier is encoded
      const accountId = userEmail || userName;
      const photoUrl = `${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(accountId)}`;
      return (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '2px' }}>
          <img
            src={photoUrl}
            alt={userName}
            style={{
              width: '24px',
              height: '24px',
              borderRadius: '50%',
              objectFit: 'cover'
            }}
            onError={(e: any) => {
              // Fallback if no photo exists
              e.target.src = "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/persona-default.png";
            }}
          />
          <span style={{ whiteSpace: 'nowrap' }}>{userName}</span>
        </div>
      );
    };
    // --- HELPER: RENDER SINGLE ROW (Encapsulates the old render logic) ---
    const renderRow = (item: any) => {
      return (
        <tr key={item.Id} className={this.state.selectedItems.indexOf(item.Id) > -1 ? styles.selectedRow : ''}>
          <td style={{ textAlign: 'center' }}>
            <input
              type="checkbox"
              checked={this.state.selectedItems.indexOf(item.Id) > -1}
              onChange={(e) => this.handleCheckboxChange(item.Id, e.target.checked)}
            />
          </td>
          {listCols.map((fName, colIndex) => {
            if (fName === 'ContentType' || fName === 'ContentTypeId') return null;

            // 1. HANDLE ATTACHMENTS ICON
            if (fName === 'Attachments') {
              const hasAttachments = item.AttachmentFiles && item.AttachmentFiles.results && item.AttachmentFiles.results.length > 0;
              const hasAttachmentsDirect = Array.isArray(item.AttachmentFiles) && item.AttachmentFiles.length > 0;
              return (
                <td key={fName} style={{ textAlign: 'center' }}>
                  {(hasAttachments || hasAttachmentsDirect || item.Attachments) ? <Icons.Clip /> : ''}
                </td>
              );
            }

            const fDef = fields.find((f: any) => f.InternalName === fName);
            const raw = item[fName];
            let val: any = '';

            // HELPER: Extract text from objects
            const getTextFromObj = (obj: any): string => {
              if (!obj) return '';
              if (typeof obj !== 'object') return String(obj);
              return obj.lookupValue || obj.LookupValue || obj.title || obj.Title || obj.Name || obj.name || obj.Description || obj.Url || obj.value || obj.Value || (obj.lookupId ? String(obj.lookupId) : '') || (obj.Id ? String(obj.Id) : '');
            };

            // 2. DATA EXTRACTION
            const isAuthorOrEditor = fName === 'Author' || fName === 'Editor';
            const isCreatedOrModified = fName === 'Created' || fName === 'Modified';
            const isUserField = (fDef && (fDef.TypeAsString === 'User' || fDef.TypeAsString === 'UserMulti')) || isAuthorOrEditor;
            const isDateField = (fDef && fDef.TypeAsString === 'DateTime') || isCreatedOrModified;

            if (raw === null || raw === undefined) {
              val = '';
            } else if (fDef && fDef.TypeAsString === 'Note') {
              const cleanText = String(raw || '').replace(/<[^>]+>/g, '');
              const displayVal = cleanText.length > 50 ? cleanText.substring(0, 50) + "..." : cleanText;
              val = <div title={cleanText} style={{ fontSize: '12px', lineHeight: '1.4', color: '#444' }}>{displayVal}</div>;
            } else if (isUserField) {
              if (typeof raw === 'object') {
                const userArr = Array.isArray(raw) ? raw : (raw.results || [raw]);
                val = <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>{userArr.map((r: any) => renderUser(r))}</div>;
              } else {
                val = renderUser({ Title: String(raw) });
              }
            } else if (isDateField) {
              val = this.formatDate(raw);
            } else if (fDef && (fDef.TypeAsString === 'LookupMulti' || fDef.TypeAsString === 'MultiChoice')) {
              val = Array.isArray(raw) ? raw : (raw.results || []);
            } else {
              if (typeof raw === 'object' && !React.isValidElement(raw)) {
                val = getTextFromObj(raw);
              } else {
                val = typeof raw === 'boolean' ? (raw ? 'Yes' : 'No') : String(raw);
              }
            }

            // 3. FORMATTING AND RENDERING
            const formattingConfig = (this.props as any).formattingConfig || {};
            const fieldFormatting = formattingConfig[fName];

            return (
              <td key={fName}>
                {(() => {
                  // A. HELPER: Generate Style
                  const getPillStyle = (text: string): React.CSSProperties => {
                    if (fieldFormatting && fieldFormatting.type === 'choice' && fieldFormatting.choiceConfig) {
                      const color = fieldFormatting.choiceConfig[text];
                      if (color && color !== '#ffffff') {
                        return { color: color, border: `1px solid ${color}`, borderRadius: '4px', padding: '2px 8px', fontWeight: 600, display: 'inline-block', fontSize: '11px' };
                      }
                    }
                    return {};
                  };

                  // B. RENDER MULTI-VALUE ARRAY
                  if (Array.isArray(val)) {
                    return (
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
                        {val.map((v: any, vIdx: number) => {
                          const text = getTextFromObj(v);
                          const style = getPillStyle(text);
                          return style.color ? <span key={vIdx} style={style}>{text}</span> : <span key={vIdx}>{text}{vIdx < val.length - 1 ? ', ' : ''}</span>;
                        })}
                      </div>
                    );
                  }

                  // C. RENDER DATE RULES
                  let dateStyle: React.CSSProperties = {};
                  if (fieldFormatting && fieldFormatting.type === 'date' && fieldFormatting.dateRules && raw) {
                    const dateVal = new Date(raw);
                    const today = new Date();
                    today.setHours(0, 0, 0, 0);
                    fieldFormatting.dateRules.forEach((rule: any) => {
                      const diffDays = Math.ceil((dateVal.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
                      let match = false;
                      if (rule.operator === 'gt_today' && dateVal > today) match = true;
                      else if (rule.operator === 'lt_today' && dateVal < today) match = true;
                      else if (rule.operator === 'within_days_before' && diffDays < 0 && Math.abs(diffDays) <= (rule.days || 0)) match = true;
                      else if (rule.operator === 'within_days_after' && diffDays > 0 && diffDays <= (rule.days || 0)) match = true;
                      if (match) dateStyle = { color: rule.color, border: `1px solid ${rule.color}`, borderRadius: '4px', padding: '2px 8px', display: 'inline-block', fontWeight: 600 };
                    });
                  }

                  // D. RENDER STATUS TAGS
                  const isStatus = fName.toLowerCase().indexOf('status') !== -1;
                  if (isStatus && !fieldFormatting) {
                    return <span className={`${styles.tag} ${this.getTagClass(String(val))}`}>{val}</span>;
                  }

                  // E. RENDER HYPERLINKS
                  if ((fDef && fDef.TypeAsString === 'URL') || (typeof val === 'string' && val.indexOf('http') === 0)) {
                    const urlHref = (raw && raw.Url) ? raw.Url : val;
                    return <a href={urlHref} target="_blank" data-interception="off" style={{ color: '#0078d4' }}>{val}</a>;
                  }

                  // F. FINAL OUTPUT
                  const choiceStyle = getPillStyle(String(val));
                  const finalStyle = dateStyle.color ? dateStyle : choiceStyle;
                  return finalStyle.color ? <span style={finalStyle}>{val}</span> : val;
                })()}
              </td>
            );
          })}
          <td>
            <div className={styles.rowActions}>
              <button className={styles.btn} title="View" onClick={() => { this.setState({ mode: 'view', itemId: item.Id }); void this.loadItemData(item.Id); }}>
                <Icons.View />
              </button>
              {this.state.canEdit && !this.props.overrideEdit && (
                <button className={styles.btn} title="Edit" onClick={() => { this.setState({ mode: 'edit', itemId: item.Id }); void this.loadItemData(item.Id); }}>
                  <Icons.Edit />
                </button>
              )}
              {this.state.canDelete && !this.props.overrideDelete && (
                <button className={`${styles.btn} ${styles.btnDanger}`} title="Delete" onClick={() => this.handleDelete(item.Id)}>
                  <Icons.Delete />
                </button>
              )}
              <button className={styles.btn} title="Share Link" onClick={() => this.handleShare(item.Id ?? 0)}>
                <i className="ms-Icon ms-Icon--Share" aria-hidden="true" style={{ fontSize: '16px' }}></i>
              </button>
            </div>
          </td>
        </tr>
      );
    };

    let activeColor = this.props.themeColor;
    if (activeColor === 'siteTheme' || !activeColor) {
      try {
        const themeState = (window as any).__themeState__;
        activeColor = (themeState && themeState.theme)
          ? themeState.theme.themePrimary
          : '#0078d4'; // Fallback
      } catch (e: any) {
        activeColor = '#0078d4';
      }
    }
    return (
      <div className={styles.crudUi} style={{ ['--accent' as any]: activeColor, ['--accent-hover' as any]: activeColor } as React.CSSProperties}>
        {isList && (
          <div className={styles.card}>
            <div className={styles.header}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 5, flex: 1 }}>
                <h1 style={{ margin: 0 }}>
                  {(this.props as any).listPageTitle || this.props.selectedList}
                </h1>
              </div>
              <div className={styles.searchContainer} style={{ flex: '0 1 450px', margin: '0 20px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <input
                  placeholder="Search all columns..."
                  value={searchText}
                  onChange={(e) => this.setState({ searchText: e.target.value })}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      this.setState({ page: 1 }, () => void this.loadItems('search'));
                    }
                  }}
                  style={{ flex: 1, padding: '6px 10px', borderRadius: '4px', border: '1px solid #ccc', height: '32px', boxSizing: 'border-box' }}
                />
                {searchText && (
                  <button
                    type="button"
                    onClick={this.clearFilters}
                    style={{ background: 'transparent', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', padding: '4px' }}
                    title="Clear Search"
                  >
                    <Icons.Clear />
                  </button>
                )}
                <button
                  type="button"
                  className={`${styles.btn} ${styles.btnPrimary}`}
                  onClick={() => this.setState({ page: 1 }, () => void this.loadItems('search'))}
                  style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 12px', height: '32px' }}
                  title="Search"
                >
                  <Icons.Search />
                </button>
              </div>
              <div className={styles.actions}>
                <button className={styles.btn} onClick={this.handleExport} title="Export">
                  <Icons.Export />
                </button>
                {/* --- BULK DELETE BUTTON (Visible only when items are selected) --- */}
                {this.state.selectedItems.length > 0 && this.state.canDelete && !this.props.overrideDelete && (
                  <button
                    className={`${styles.btn} ${styles.btnDanger}`}
                    onClick={this.handleBulkDelete}
                    title="Delete Selected Items"
                    style={{ marginRight: '10px', display: 'flex', alignItems: 'center', gap: '5px' }}
                  >
                    <Icons.Delete />
                    <span style={{ fontWeight: 600 }}>Delete ({this.state.selectedItems.length})</span>
                  </button>
                )}
                {this.state.selectedItems.length === 1 && (
                  <button className={styles.btn} title="View" onClick={() => { this.setState({ mode: 'view', itemId: this.state.selectedItems[0] }); void this.loadItemData(this.state.selectedItems[0]); }}><Icons.View /></button>
                )}

                {this.state.selectedItems.length === 1 && this.state.canEdit && !this.props.overrideEdit && (
                  <button className={styles.btn} title="Edit" onClick={() => { this.setState({ mode: 'edit', itemId: this.state.selectedItems[0] }); void this.loadItemData(this.state.selectedItems[0]); }}><Icons.Edit /></button>
                )}

                {/*  HIDE ADD BUTTON IF NO PERMISSION */}
                {this.state.canAdd && !this.props.overrideAdd && (
                  <button className={`${styles.btn} ${styles.btnPrimary}`} onClick={() => this.setState({
                    mode: 'add', formData: {},
                    itemId: undefined, childItems: {}, attachmentsNew: [], existingAttachments: []
                  })} title="Add New">
                    <Icons.Plus />
                  </button>
                )}
                {(this.state.availableViews.length > 0 || !this.state.canSeeDefaultView) && (
                  <div style={{ marginRight: 8 }}>
                    <select
                      value={this.state.currentViewId}
                      onChange={(e) => this.handleViewChange(e.target.value)}
                      style={{ padding: '0 8px', borderRadius: '4px', border: '1px solid #ccc', height: '32px' }}
                    >
                      {/* CONDITIONALLY RENDER DEFAULT OPTION */}
                      {this.state.canSeeDefaultView && <option value="">All Items</option>}
                      {this.state.availableViews.map((v: any) => (
                        <option key={v.id} value={v.id}>{v.title}</option>
                      ))}
                    </select>
                  </div>
                )}
                {/* {this.state.selectedItems.length > 0 && (
                  <button
                    className={`${styles.btn} ${styles.btnPrimary}`}
                    onClick={() => this.setState({ isBulkEditOpen: true })}
                  >
                    Bulk Edit ({this.state.selectedItems.length})
                  </button>
                )} */}
              </div>
            </div>
            <div className={styles.tableWrap}>
              <table>
                <thead>
                  <tr>
                    <th style={{ width: 40, textAlign: 'center' }}>
                      <input
                        type="checkbox"
                        onChange={(e) => this.handleSelectAll(e.target.checked)}
                        checked={items.length > 0 && this.state.selectedItems.length === items.length}
                      />
                    </th>
                    {listCols.map((fName, idx) => {
                      if (fName === 'ContentType' || fName === 'ContentTypeId') return null;
                      const fDef = fields.find(f => f.InternalName === fName);
                      const title = fDef ? fDef.Title : fName;
                      const isFilterEnabled = hasCustomFilters ? filterConfig[fName] === true : true;
                      const rawFilter = filters[fName];
                      const filterValue = (rawFilter && typeof rawFilter === 'object' && (rawFilter as any).value) ? (rawFilter as any).value : (rawFilter || '');
                      const isDateTime = fDef && fDef.TypeAsString === 'DateTime';
                      const isAttachment = fName === 'Attachments';
                      const isBooleanFilter = fDef && fDef.TypeAsString === 'Boolean';
                      const isHyperlinkFilter = fDef && fDef.TypeAsString === 'URL';
                      const showInput = !isAttachment && !isDateTime && isFilterEnabled && !isBooleanFilter && !isHyperlinkFilter;
                      return (
                        <th key={fName}>
                          <div style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }} onClick={() => this.handleSort(fName)}>
                            {title}
                            {this.state.sortField === fName ? (this.state.sortDirection === 'asc' ? <Icons.SortAsc /> : <Icons.SortDesc />) : null}
                          </div>
                          {showInput && (
                            <input
                              className={styles.filterInput}
                              placeholder={`Filter ${title}...`}
                              value={filterValue}
                              onChange={(e) => this.handleFilterChange(fName, e.target.value)}
                            />
                          )}
                        </th>
                      );
                    })}
                    <th style={{ width: 140 }}>Actions</th>
                  </tr>
                </thead>

                <tbody>
                  {items.length === 0 ? (
                    <tr><td colSpan={listCols.length + 2} className={styles.emptyState}>No items match your filters.</td></tr>
                  ) : (
                    (() => {
                      // --- LOGIC: BUILD FLAT ARRAY OF ROWS (No Fragments, No Divs) ---
                      const rows: JSX.Element[] = [];
                      const groupField = (this.props as any).listGroupByField;

                      if (groupField && currentItems.length > 0) {
                        // 1. Group the items
                        const groups: { [key: string]: any[] } = {};
                        currentItems.forEach(item => {
                          const val = item[groupField] || 'Uncategorized';
                          const key = typeof val === 'object' ? (val.Title || val) : String(val);
                          if (!groups[key]) groups[key] = [];
                          groups[key].push(item);
                        });

                        // 2. Iterate Groups and Push to Array
                        Object.keys(groups).sort().forEach(gKey => {
                          // A. Push Header Row
                          rows.push(
                            <tr key={`group_${gKey}`} style={{ backgroundColor: '#f3f2f1', fontWeight: 600 }}>
                              <td colSpan={listCols.length + 2} style={{ padding: '10px 15px', borderBottom: '1px solid #e1dfdd' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                  <span style={{ fontSize: '14px' }}>📂</span>
                                  <span style={{ color: '#333' }}>{gKey}</span>
                                  <span style={{ background: '#fff', padding: '2px 8px', borderRadius: '10px', fontSize: '11px', border: '1px solid #ccc', color: '#666' }}>
                                    {groups[gKey].length}
                                  </span>
                                </div>
                              </td>
                            </tr>
                          );

                          // B. Push Data Rows for this Group
                          groups[gKey].forEach(item => {
                            rows.push(renderRow(item));
                          });
                        });
                      } else {
                        // --- STANDARD RENDER (No Grouping) ---
                        currentItems.forEach(item => {
                          rows.push(renderRow(item));
                        });
                      }

                      return rows;
                    })()
                  )}
                </tbody>
              </table>
            </div>
            {this.props.isLargeList ? this.renderPagination_Above5K() : this.renderPagination()}
          </div>
        )}
        {!isList && (
          <div className={styles.card}>
            <div className={styles.formHeader}>
              <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 600 }}>
                {(() => {
                  const listPrefix = this.props.selectedList;
                  // Access custom title props passed from the Web Part
                  const customTitles = this.props as any;
                  if (mode === 'add') {
                    return customTitles.addPageTitle || `${listPrefix} - Add Item`;
                  }
                  if (mode === 'edit') {
                    return customTitles.editPageTitle || `${listPrefix} - Edit Item`;
                  }
                  if (mode === 'view') {
                    return customTitles.viewPageTitle || `${listPrefix} - View Item`;
                  }
                  return listPrefix;
                })()}
              </h2>
              <div className={styles.actions}>
                {/* NEW: RENDER CUSTOM BUTTONS */}
                {customActions.map((action, idx) => (
                  <button
                    key={idx}
                    className={styles.btn}
                    onClick={() => this.handleCustomAction(action)}
                    title={action.title}
                    style={{ display: 'flex', gap: 6 }}
                  >
                    {/* Optional Icon if you implement icon lookup, otherwise just text */}
                    <span style={{ fontWeight: 600 }}>{action.title}</span>
                  </button>
                ))}
                {mode !== 'add' && this.state.enableVersioning && this.state.itemId && (
                  <button className={styles.btn} title="Version History" onClick={this.openVersionHistory}>
                    <Icons.History />
                  </button>
                )}
                {mode !== 'add' && this.state.itemId && (
                  <button className={styles.btn} type="button" onClick={() => this.handleShare(this.state.itemId ?? 0)} title="Copy Link to this Item">
                    <i className="ms-Icon ms-Icon--Share" aria-hidden="true"></i>
                  </button>
                )}
                {/*  HIDE EDIT BUTTON IN VIEW MODE IF NO PERMISSION */}
                {mode === 'view' && this.state.canEdit && !this.props.overrideEdit && (
                  <button className={styles.btn} title="Edit" onClick={() => this.setState({ mode: 'edit' })}>
                    <Icons.Edit />
                  </button>
                )}
                {mode !== 'add' && this.state.itemId && this.state.canDelete && !this.props.overrideDelete && (
                  <button
                    className={`${styles.btn} ${styles.btnDanger}`}
                    title="Delete Item"
                    onClick={() => this.handleDelete(this.state.itemId!)}
                  >
                    <Icons.Delete />
                  </button>
                )}
                {(mode === 'edit' || mode === 'view') && (
                  <button
                    type="button"
                    className={styles.btn} // Use existing button style
                    title="Refresh Item"
                    onClick={() => {
                      if (this.state.itemId) {
                        //  Show Loading
                        this.setState({ loading: true });
                        //  Reload Data
                        void this.loadItemData(this.state.itemId)
                          .then(() => {
                            //  Re-Apply URL Params (Optional: if you want URL params to override again)
                            if (mode === 'edit') this.applyUrlParameters();
                            this.setState({ loading: false });
                          })
                          .catch(err => {
                            this.setState({ loading: false });
                          });
                      }
                    }}
                  >
                    <Icons.Refresh />
                  </button>
                )}
                <button className={styles.btn} title="Back" onClick={() => this.setState({
                  mode: 'list',
                  itemId: undefined, formData: {}, childItems: {}, existingAttachments: []
                })}>
                  <Icons.Back />
                </button>
              </div>
            </div>
            <form className={styles.formBody} onSubmit={this.handleSubmit}>
              {(() => {
                const sections = this.props.formSections || [];
                const layout = (this.props.sectionLayout as any) || 'stacked';
                const { activeSectionIndex, mode, formData } = this.state;
                const activeColor = this.props.themeColor || '#0078d4';
                // layoutClass is grid1 (Single Column) or grid2 (Two Columns)
                const layoutClass = this.props.formLayout === 'single' ? styles.grid1 : styles.grid2;
                const usedFieldKeys = new Set<string>();
                // 1. HELPER: Render Audit Information (Created/Modified)
                const renderAuditRow = () => {
                  if ((mode !== 'edit' && mode !== 'view') || !this.state.itemId) return null;
                  return (
                    <div style={{ gridColumn: '1 / -1', borderTop: '1px solid #eee', marginTop: '20px', paddingTop: '15px' }}>
                      <div className={layoutClass}>
                        <div className={styles.field}>
                          <label style={{ fontWeight: 800 }}>Created</label>
                          <div className={styles.viewField}>
                            {this.formatDate_form(formData['Created'])}
                            <span style={{ margin: '0 8px', color: '#999' }}>by</span>
                            {formData['Author'] ? renderUser(formData['Author']) : '-'}
                          </div>
                        </div>
                        <div className={styles.field}>
                          <label style={{ fontWeight: 800 }}>Modified</label>
                          <div className={styles.viewField}>
                            {this.formatDate_form(formData['Modified'])}
                            <span style={{ margin: '0 8px', color: '#999' }}>by</span>
                            {formData['Editor'] ? renderUser(formData['Editor']) : '-'}
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                };
                // 2. ORPHAN & INDEX LOGIC
                sections.forEach(s => s.fields.forEach(f => usedFieldKeys.add(f)));
                const orphans = formFields.filter(f => !usedFieldKeys.has(f.InternalName));
                const hasOrphans = orphans.length > 0;
                const totalStepsCount = hasOrphans ? sections.length + 1 : sections.length;
                const finalIndex = totalStepsCount - 1;
                // 3. NAVIGATION RENDERERS
                const renderTabs = () => (
                  <div style={{ display: 'flex', borderBottom: '1px solid #ccc', marginBottom: 25, overflowX: 'auto', gap: '10px' }}>
                    {sections.map((s, idx) => (
                      <div key={s.id} onClick={() => this.setState({ activeSectionIndex: idx })}
                        style={{
                          padding: '10px 20px', cursor: 'pointer', fontWeight: 800, fontSize: '14px', whiteSpace: 'nowrap',
                          borderBottom: activeSectionIndex === idx ? `3px solid ${activeColor}` : '3px solid transparent',
                          color: activeSectionIndex === idx ? activeColor : '#666'
                        }}> {s.title} </div>
                    ))}
                    {hasOrphans && (
                      <div onClick={() => this.setState({ activeSectionIndex: sections.length })}
                        style={{
                          padding: '10px 20px', cursor: 'pointer', fontWeight: 800, fontSize: '14px',
                          borderBottom: activeSectionIndex === sections.length ? `3px solid ${activeColor}` : '3px solid transparent',
                          color: activeSectionIndex === sections.length ? activeColor : '#666'
                        }}> General / Others </div>
                    )}
                  </div>
                );
                const renderWizardHeader = () => (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 30, flexWrap: 'wrap', padding: '0 5px' }}>
                    {sections.map((s, idx) => {
                      const isActive = activeSectionIndex === idx;
                      return (
                        <div key={s.id} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                          <div style={{
                            fontSize: '14px', fontWeight: isActive ? 700 : 400,
                            color: isActive ? activeColor : '#999',
                            padding: '5px 10px', borderRadius: '4px', background: isActive ? `${activeColor}10` : 'transparent'
                          }}>
                            {idx + 1}. {s.title}
                          </div>
                          {(idx < sections.length - 1 || hasOrphans) && <span style={{ color: '#ccc', fontWeight: 'bold' }}>→</span>}
                        </div>
                      );
                    })}
                    {hasOrphans && (
                      <div style={{
                        fontSize: '14px', fontWeight: activeSectionIndex === sections.length ? 700 : 400,
                        color: activeSectionIndex === sections.length ? activeColor : '#999',
                        padding: '5px 10px', borderRadius: '4px', background: activeSectionIndex === sections.length ? `${activeColor}10` : 'transparent'
                      }}>
                        {sections.length + 1}. General / Others
                      </div>
                    )}
                  </div>
                );
                // 4. MAIN RENDER LOGIC
                if (layout === 'none' || sections.length === 0) {
                  return (
                    <div className={layoutClass}>
                      {formFields.map(f => this.renderFormField(f, mode === 'view', readOnlyKeys.indexOf(f.InternalName) > -1))}
                      {this.props.childConfigs && this.props.childConfigs.map(config => this.renderChildSection(config))}
                      {renderAuditRow()}
                    </div>
                  );
                }
                return (
                  <div>
                    {layout === 'tabs' && renderTabs()}
                    {layout === 'wizard' && renderWizardHeader()}
                    {/* Configured Sections */}
                    {sections.map((section, idx) => {
                      if (layout !== 'stacked' && activeSectionIndex !== idx) return null;
                      const sectionFields = formFields.filter(f => section.fields.indexOf(f.InternalName) > -1);
                      const isFinalStepInUI = idx === finalIndex;
                      return (
                        <div key={section.id} className={styles.formSection} style={{ border: `2px solid ${activeColor}`, marginBottom: '20px' }}>
                          <h3 style={{ color: '#333' }}>{section.title}</h3>
                          <div className={layoutClass}>
                            {sectionFields.map(f => this.renderFormField(f, mode === 'view', readOnlyKeys.indexOf(f.InternalName) > -1))}
                          </div>
                          {isFinalStepInUI && this.props.childConfigs && this.props.childConfigs.map(config => this.renderChildSection(config))}
                          {isFinalStepInUI && renderAuditRow()}
                          {layout === 'wizard' && (
                            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 25, paddingTop: 20, borderTop: '1px solid #eee' }}>
                              <button type="button" className={styles.btn} disabled={activeSectionIndex === 0} onClick={() => this.setState({ activeSectionIndex: activeSectionIndex - 1 })}>Previous</button>
                              {activeSectionIndex < finalIndex && (
                                <button type="button" className={`${styles.btn} ${styles.btnPrimary}`} onClick={() => this.setState({ activeSectionIndex: activeSectionIndex + 1 })}>Next Step</button>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                    {/* Orphan Section */}
                    {hasOrphans && (layout === 'stacked' || activeSectionIndex === sections.length) && (
                      <div className={styles.formSection} style={{ border: '2px solid #ccc' }}>
                        <h3 style={{ color: '#666' }}>General / Others</h3>
                        <div className={layoutClass}>
                          {orphans.map(f => this.renderFormField(f, mode === 'view', readOnlyKeys.indexOf(f.InternalName) > -1))}
                        </div>
                        {this.props.childConfigs && this.props.childConfigs.map(config => this.renderChildSection(config))}
                        {renderAuditRow()}
                        {layout === 'wizard' && (
                          <div style={{ display: 'flex', justifyContent: 'flex-start', marginTop: 25, paddingTop: 20, borderTop: '1px solid #eee' }}>
                            <button type="button" className={styles.btn} onClick={() => this.setState({ activeSectionIndex: sections.length - 1 })}>Previous</button>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                );
              })()}

              {/* 5. SUBMIT ACTIONS: Visible only on the final step or in continuous layouts */}
              {(() => {
                const layout = this.props.sectionLayout || 'stacked';
                const sections = this.props.formSections || [];
                const usedKeys = new Set((this.props.formSections || []).reduce<string[]>((acc, s) => acc.concat(s.fields), []));
                const hasOrphans = formFields.some(f => !usedKeys.has(f.InternalName));
                const finalIndex = hasOrphans ? sections.length : sections.length - 1;
                const isLastStep = layout === 'none' || layout === 'stacked' || this.state.activeSectionIndex === finalIndex;
                if (this.state.mode !== 'view' && isLastStep) {
                  return (
                    <div className={styles.formActions} style={{ marginTop: '20px', display: 'flex', gap: '10px' }}>
                      <button type="button" className={styles.btn} onClick={() => this.setState({ formData: {}, activeSectionIndex: 0 })}>
                        <Icons.Reset /> Reset
                      </button>
                      <button type="submit" className={`${styles.btn} ${styles.btnPrimary}`} disabled={isSaveDisabled}>
                        <Icons.Save /> Save
                      </button>
                    </div>
                  );
                }
              })()}
            </form>
          </div>
        )}
        {this.state.loading && (
          <div style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(255,255,255,0.7)', zIndex: 99, display: 'flex', justifyContent: 'center', alignItems: 'center' }}>
            Loading...
          </div>
        )}
        {/* NEW: SLIDER PANEL */}
        <Panel
          isOpen={this.state.isPanelOpen}
          onDismiss={() => this.setState({ isPanelOpen: false })}
          //  CHANGE WIDTH: Use Custom type to allow wider panels (e.g., 70% of screen)
          type={PanelType.custom}
          customWidth="70%"
          headerText={this.state.panelTitle}
          closeButtonAriaLabel="Close"
          isLightDismiss={true}
        >
          <div style={{ height: 'calc(100vh - 100px)', width: '100%', overflow: 'hidden' }}>
            {this.state.isPanelOpen && this.state.panelUrl && (
              <iframe
                src={this.state.panelUrl}
                width="100%"
                height="100%"
                frameBorder="0"
                style={{ border: 'none' }}
                // ISOLATE CONTENT: Inject CSS when iframe loads
                onLoad={(e) => {
                  try {
                    const iframe = e.target as HTMLIFrameElement;
                    // Check access (only works on same-domain/SharePoint pages)
                    // Using optional chaining means doc could be undefined
                    const doc = iframe.contentDocument || iframe.contentWindow?.document;

                    // FIX: Wrap in a check to ensure doc is not null or undefined before use
                    if (doc) {
                      // Create Style Tag
                      const style = doc.createElement('style');
                      style.textContent = `
    #spPageChromeAppDiv, .sp-App-body {
        margin-left: 0px !important;
    }
    .CanvasSection-col {
        padding-left: 0px !important; 
        padding-right: 0px !important;
    }
    .crudUi .card {
        margin: 0px !important;
        background-color: #ffffff !important;
        border: none !important; /* Optional: Looks cleaner in panel */
        box-shadow: none !important; /* Optional: Looks cleaner in panel */
    }
    @media screen and (min-width: 641px) {
        .CanvasZone {
            padding: 0 0px !important;
        }
    }
    /* --- STANDARD CLEANUP (Hide Headers/Nav) --- */
    .spPageCanvasContent {
        padding: 0px !important;
        margin: 0px !important;
        max-width: 100% !important;
    }
    .site-menu,
    #SuiteNavPlaceHolder, 
    #spCommandBar, 
    .sp-appBar, 
    div[data-automation-id="pageHeader"],
    #spLeftNav,
    #spSiteHeader,
    #sp-appBar-placeholder { 
        display: none !important; 
    } 
  `;

                      // Safe to append because of the 'if (doc)' guard
                      doc.head.appendChild(style);
                    }
                  } catch (error: any) {
                    void LoggerService.log(
                      'PowerForm - IframeLoad',
                      'Low',
                      'N/A',
                      error.message || JSON.stringify(error)
                    );
                  }
                }}
              />
            )}
          </div>
        </Panel>

        {this.props.childConfigs && this.props.childConfigs.length > 0 && (
          <Panel
            isOpen={this.state.isChildPanelOpen}
            onDismiss={() => this.setState({ isChildPanelOpen: false })}
            headerText={this.state.activeChildConfig ? `Add ${this.state.activeChildConfig.title}` : 'Add Item'}
            type={PanelType.medium}
          >
            {this.state.activeChildConfig && (
              <div>
                {/* Simple Form Renderer for Child Item */}
                <div className={styles.formBody}>
                  {(() => {
                    const config = this.state.activeChildConfig;
                    const fields = this.state.childFieldsCache[config.childListTitle] || [];
                    // Only show fields defined in config
                    const visibleFields = fields.filter((f: any) => config.visibleFields.indexOf(f.InternalName) > -1);

                    // Render Inputs
                    return visibleFields.map((f: any) => {
                      const listKey = config.childListTitle;
                      const items = this.state.childItems[listKey] || [];

                      return (
                        <div key={f.InternalName} className={styles.field}>
                          <label>{f.Title}</label>
                          {/* Re-using grid input logic for simplicity in this modal */}
                          {this.state.activeChildItemIndex > -1 ? (
                            this.renderChildInput(listKey, this.state.activeChildItemIndex, f, items[this.state.activeChildItemIndex][f.EntityPropertyName])
                          ) : (
                            <div>Please add a row in Grid mode first, then edit. (Form Mode Pending full implementation)</div>
                          )}
                        </div>
                      );
                    });
                  })()}
                </div>

                <div style={{ marginTop: 20 }}>
                  <button className={`${styles.btn} ${styles.btnPrimary}`} onClick={() => this.setState({ isChildPanelOpen: false })}>
                    Done
                  </button>
                </div>
              </div>
            )}
          </Panel>
        )}
      </div>
    );
  }
}