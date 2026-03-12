
import * as React from 'react';
import * as ReactDom from 'react-dom';
import Swal from 'sweetalert2';
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/security";
import "@pnp/sp/site-users";
import "@pnp/sp/site-groups";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  IPropertyPaneCustomFieldProps,
  IPropertyPanePage,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel,
  PropertyPaneToggle,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { PropertyPaneFieldType } from '@microsoft/sp-webpart-base';
// import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { SPHttpClient } from '@microsoft/sp-http';
import PowerForm from './components/PowerForm';
import { IPowerFormProps } from './components/IPowerFormProps';
import { FieldOrderRenderer } from '../../Common/Controls/FieldOrderRenderer';
import { CommonService } from '../../Common/Services/CommonService';
import { ValidationEditor } from './components/ValidationEditor';
import { IValidationConfig, ICustomValidationRule } from './components/ICustomValidation';
import { ICustomAction } from './components/ICustomAction';
import { CustomActionEditor } from './components/CustomActionEditor';
import { FieldPermissionEditor } from './components/FieldPermissionEditor';
import { ViewEditor, IViewConfig } from './components/ViewEditor';
import { LoggerService } from './components/LoggerService';
import { LogViewer } from './components/LogViewer';
import { IFormSection } from './components/IFormSection';
import { SectionEditor } from './components/SectionEditor';
import { FormattingEditor } from './components/FormattingEditor';
import { RepeaterConfigEditor } from './components/RepeaterConfigEditor';
import { IRepeaterColumn, IRepeaterConfig } from './components/IPowerFormProps';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';


export function PropertyPaneCustomField(properties: IPropertyPaneCustomFieldProps): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty: properties.key || 'CustomField',
    properties: properties
  };
}

// --- ICONS DEFINITION ---
// Simple SVG icons used within the Property Pane buttons (Lock, Unlock, Filter, etc.)
const iconStyle = { width: 16, height: 16, fill: 'none', stroke: 'currentColor', strokeWidth: 2, verticalAlign: 'middle' };
const Icons = {
  Lock: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("rect", { x: "3", y: "11", width: "18", height: "11", rx: "2", ry: "2" }), React.createElement("path", { d: "M7 11V7a5 5 0 0 1 10 0v4" })),
  Unlock: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("rect", { x: "3", y: "11", width: "18", height: "11", rx: "2", ry: "2" }), React.createElement("path", { d: "M7 11V7a5 5 0 0 1 9.9-1" })),
  Filter: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("polygon", { points: "22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" })),
  FilterOff: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("polygon", { points: "22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3", strokeDasharray: "4 2" })),
  Cascade: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("path", { d: "M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" }), React.createElement("path", { d: "M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" })),
  Autocomplete: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("path", { d: "M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" }), React.createElement("path", { d: "M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" })),
  Validation: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("path", { d: "M9 11l3 3L22 4" }), React.createElement("path", { d: "M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" })),
  Group: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("path", { d: "M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-5.5-2.5l7.51-3.22-3.22-7.51-7.51 3.22 3.22 7.51zm3.22-4.29c.59.59 1.55.59 2.14 0 .59-.59.59-1.55 0-2.14-.59-.59-1.55-.59-2.14 0-.59.59-.59 1.55 0 2.14z" })),
  Grid: () => React.createElement('svg', { viewBox: "0 0 2048 2048", width: 16, height: 16, fill: "currentColor" }, React.createElement('path', { d: "M1024 896v1024H0V896h1024zm-128 896V1024H128v768h768zM1024 0v896H0V0h1024zm-128 128H128v640h768V128zm1152 768v1024h-1024V896h1024zm-128 896V1024h-768v768h768zM2048 0v896h-1024V0h1024zm-128 128h-768v640h768V128z" })),
  Permissions: () => React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("path", { d: "M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm-2 16l-4-4 1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z" }))

};
/** * Helper: Shallow-clone a map of strings to numbers.
 * Used for cloning Field Order maps.
 */
function cloneMap(source: { [key: string]: number } | undefined): { [key: string]: number } {
  let dest: { [key: string]: number } = {};
  if (!source) { return dest; }
  for (let k in source) {
    if (Object.prototype.hasOwnProperty.call(source, k)) { dest[k] = source[k]; }
  }
  return dest;
}
/** * Helper: Shallow-clone validationConfig.
 * Essential for immutability when updating React/SPFx state.
 */
function cloneConfig(source: IValidationConfig | undefined): IValidationConfig {
  let dest: IValidationConfig = {};
  if (!source) { return dest; }
  for (let f in source) {
    if (Object.prototype.hasOwnProperty.call(source, f)) { dest[f] = source[f]; }
  }
  return dest;
}

/** ES5-friendly Set→Array helper */
function setToArray(s: Set<any>): string[] {
  let out: string[] = [];
  s.forEach(function (v) { out.push(v); });
  return out;
}
// --- INTERFACES FOR CONFIGURATION OBJECTS ---
/** Choice Formatting Config */
export interface IChoiceFormatConfig {
  [choiceValue: string]: string; // choiceValue -> hex color
}
/** Date Formatting Rule */
export interface IDateFormatRule {
  condition: 'past' | 'future' | 'past_n' | 'future_n' | 'today';
  days?: number;
  color: string;
}
/** Validation Configuration for fields */
export interface IFieldFormatting {
  type: 'choice' | 'date';
  choiceConfig?: IChoiceFormatConfig;
  dateRules?: IDateFormatRule[];
}
/** Configuration for field formatting (colors for choices, date rules, etc.) */
export interface IFormattingConfig {
  [fieldKey: string]: {
    type: 'choice' | 'date';
    choiceConfig?: IChoiceFormatConfig;
    dateRules?: IDateFormatRule[];
  };
}
/** Defines a mapping between a column in a source list and a field in the current form. */
export interface IColumnMapping {
  source: string;
  target: string;
}
/** Configuration for Autocomplete fields (Source List, Fields, Query) */
export interface IAutocompleteConfig {
  [fieldKey: string]: {
    sourceList: string;
    sourceField: string;
    sourceQuery?: string;
    additionalFields?: string[];
    columnMapping?: IColumnMapping[];
  };
}
/** Configuration for Lookup fields (Columns to display, Filters) */
export interface ILookupDisplayConfig {
  [fieldKey: string]: {
    additionalFields: string[];
    filterQuery?: string;
    columnMapping?: IColumnMapping[];
  };
}
/** Configuration for Cascading Lookups (Parent/Child relationships) */
export interface ICascadeConfig {
  [childField: string]: {
    parentField: string;   // e.g. "Country"
    foreignKey: string;    // The internal name of the Lookup column in the Child List (e.g. "CountryRef")
    filterQuery?: string;
    additionalFields?: string[];
    columnMapping?: IColumnMapping[];
  };
}
export interface IChildListConfig {
  childListTitle: string;
  foreignKeyField: string; // The lookup field in Child List pointing to Parent
  uiMode: 'row' | 'form';
  visibleFields: string[];
  title: string;
}

// --- HELPER COMPONENT: NotificationRuleEditor ---
export interface INotificationRule {
  enabled: boolean;
  condition: string;
  message: string;
  targetGroups: number[];
}


// --- HELPER COMPONENT: CheckboxListEditor ---
// An alternative to MultiSelect, renders checkboxes in a scrollable div.
// Supports a max selection limit.
class CheckboxListEditor extends React.Component<{
  label: string;
  options: { key: string; text: string }[];
  selectedKeys: string[];
  onChanged: (keys: string[]) => void;
  maxSelection?: number; // Optional limit
}, { selectedKeys: string[] }> {
  constructor(props: any) {
    super(props);
    try {
      this.state = { selectedKeys: props.selectedKeys || [] };
    } catch (error: any) {
      void LoggerService.log('CheckboxListEditor-constructor', 'High', 'Config', error.message || JSON.stringify(error));
    }
  }
  public componentWillReceiveProps(nextProps: any) {
    try {
      if (JSON.stringify(nextProps.selectedKeys) !== JSON.stringify(this.state.selectedKeys)) {
        this.setState({ selectedKeys: nextProps.selectedKeys || [] });
      }
    } catch (error: any) {
      void LoggerService.log('CheckboxListEditor-willReceiveProps', 'Low', 'Config', error.message || JSON.stringify(error));
    }
  }
  private toggleKey = (key: string) => {
    try {
      const current = [...this.state.selectedKeys];
      const idx = current.indexOf(key);
      if (idx > -1) {
        // Uncheck is always allowed
        current.splice(idx, 1);
      } else {
        // Check Limit before adding
        if (this.props.maxSelection && current.length >= this.props.maxSelection) {
          void Swal.fire({ icon: 'warning', title: 'warning', text: `You can only select up to ${this.props.maxSelection} columns.` });
          return;
        }
        current.push(key);
      }
      this.setState({ selectedKeys: current });
      this.props.onChanged(current);
    } catch (error: any) {
      void LoggerService.log('CheckboxListEditor-toggleKey', 'Medium', 'Config', error.message || JSON.stringify(error));
    }
  };
  public render() {
    return React.createElement('div', { style: { marginTop: 8 } },
      React.createElement('label', { style: { fontWeight: 600, display: 'block', marginBottom: 8 } }, this.props.label),
      // Scrollable Container
      React.createElement('div', {
        style: {
          maxHeight: '200px',
          overflowY: 'auto',
          border: '1px solid #e1dfdd',
          padding: '5px',
          background: '#ffffff'
        }
      },
        this.props.options.map((opt: any) => {
          const isChecked = this.state.selectedKeys.indexOf(opt.key) > -1;
          const isDisabled = !isChecked && this.props.maxSelection && this.state.selectedKeys.length >= this.props.maxSelection;
          return React.createElement('div', {
            key: opt.key,
            style: {
              display: 'flex',
              alignItems: 'center',
              marginBottom: 6,
              cursor: isDisabled ? 'not-allowed' : 'pointer',
              opacity: isDisabled ? 0.6 : 1
            },
            onClick: () => !isDisabled && this.toggleKey(opt.key)
          },
            React.createElement('input', {
              type: 'checkbox',
              checked: isChecked,
              disabled: isDisabled,
              onChange: () => { },
              style: { marginRight: 8, cursor: isDisabled ? 'not-allowed' : 'pointer' }
            }),
            React.createElement('span', null, opt.text)
          );
        })
      )
    );
  }
}


export class NotificationRuleEditor extends React.Component<{
  label: string;
  rules: INotificationRule[];
  siteGroups: { key: string; text: string }[];
  onSave: (rules: INotificationRule[]) => void;
  onCancel: () => void;
}, { rules: INotificationRule[] }> {
  constructor(props: any) {
    super(props);
    this.state = { rules: props.rules ? JSON.parse(JSON.stringify(props.rules)) : [] };
  }

  private addRule = () => {
    const newRule: INotificationRule = { enabled: true, condition: '', message: '', targetGroups: [] };
    this.setState({ rules: this.state.rules.concat([newRule]) });
  }

  private removeRule = (index: number) => {
    const newRules = this.state.rules.slice();
    newRules.splice(index, 1);
    this.setState({ rules: newRules });
  }

  private updateRule = (index: number, field: keyof INotificationRule, value: any) => {
    const newRules = this.state.rules.slice();
    (newRules[index] as any)[field] = value;
    this.setState({ rules: newRules });
  }

  public render() {
    return React.createElement('div', { style: { padding: '10px', background: '#f8f8f8', border: '1px solid #ddd', marginTop: '10px' } },
      React.createElement('h4', { style: { marginTop: 0 } }, this.props.label),

      this.state.rules.map((rule, idx) => {
        return React.createElement('div', { key: idx, style: { padding: '10px', border: '1px solid #ccc', background: '#fff', marginBottom: '10px', position: 'relative' } },
          // Delete Button
          React.createElement('button', {
            onClick: () => this.removeRule(idx),
            style: { position: 'absolute', right: '5px', top: '5px', color: 'red', cursor: 'pointer', border: 'none', background: 'transparent', fontWeight: 'bold' },
            title: 'Delete Rule'
          }, 'X'),

          // Active Toggle
          React.createElement('div', { style: { marginBottom: '8px' } },
            React.createElement('label', { style: { fontWeight: 600, fontSize: '12px', display: 'flex', alignItems: 'center' } },
              React.createElement('input', {
                type: 'checkbox',
                checked: rule.enabled,
                onChange: (e: any) => this.updateRule(idx, 'enabled', e.target.checked),
                style: { marginRight: '5px' }
              }),
              "Rule is Active"
            )
          ),

          // Condition Input (WITH NEW PLACEHOLDER)
          React.createElement('div', { style: { marginBottom: '8px' } },
            React.createElement('label', { style: { display: 'block', fontSize: '12px', fontWeight: 600 } }, "Condition (JS/OData)"),
            React.createElement('input', {
              style: { width: '100%', padding: '6px', boxSizing: 'border-box' },
              value: rule.condition,
              onChange: (e: any) => this.updateRule(idx, 'condition', e.target.value),
              // [UPDATED PLACEHOLDER]
              placeholder: "e.g. item.Title eq 'Test' and item.Status eq 'Not Started'"
            })
          ),

          // Message Input (WITH NEW PLACEHOLDER)
          React.createElement('div', { style: { marginBottom: '8px' } },
            React.createElement('label', { style: { display: 'block', fontSize: '12px', fontWeight: 600 } }, "Message Template *"),
            React.createElement('textarea', {
              style: { width: '100%', padding: '6px', boxSizing: 'border-box' },
              value: rule.message,
              onChange: (e: any) => this.updateRule(idx, 'message', e.target.value),
              rows: 3,
              // [UPDATED PLACEHOLDER]
              placeholder: "e.g. {ListName} {Title} updated by {loggeduser} / {currentuser}"
            })
          ),

          // Group Picker
          React.createElement('div', { style: { marginBottom: '8px' } },
            React.createElement(CheckboxListEditor, {
              label: "Target SharePoint Groups",
              options: this.props.siteGroups,
              selectedKeys: (rule.targetGroups || []).map(String),
              onChanged: (keys: string[]) => this.updateRule(idx, 'targetGroups', keys.map(k => parseInt(k, 10)))
            })
          )
        );
      }),

      React.createElement('button', {
        onClick: this.addRule,
        style: { padding: '8px 12px', background: '#eff6ff', color: '#0078d4', border: '1px solid #0078d4', cursor: 'pointer', marginBottom: '15px', width: '100%', fontWeight: 600, borderRadius: '4px' }
      }, "+ Add New Rule"),

      React.createElement('div', { style: { display: 'flex', gap: '10px' } },
        React.createElement('button', {
          onClick: () => this.props.onSave(this.state.rules),
          style: { padding: '8px 12px', background: '#107c10', color: '#fff', border: 'none', cursor: 'pointer', flex: 1, borderRadius: '4px', fontWeight: 600 }
        }, "Save Rules"),
        React.createElement('button', {
          onClick: this.props.onCancel,
          style: { padding: '8px 12px', background: '#666', color: '#fff', border: 'none', cursor: 'pointer', flex: 1, borderRadius: '4px', fontWeight: 600 }
        }, "Cancel")
      )
    );
  }
}

/** * Main WebPart Properties Interface
 * Stores all configuration state persisted by SharePoint
 */
export interface IPowerFormWebPartProps {
  selectedList: string;
  isLargeList: boolean;
  // ADD Form Config
  addVisibleFields: string[];
  addFieldOrder: { [key: string]: number };
  addReadOnlyFields: string[];
  addCustomScript: string;
  addCustomStyle: string;
  // EDIT Config
  editVisibleFields: string[];
  editFieldOrder: { [key: string]: number };
  editReadOnlyFields: string[];
  editCustomScript: string;
  editCustomStyle: string;
  // VIEW Config
  viewVisibleFields: string[];
  viewFieldOrder: { [key: string]: number };
  viewCustomScript: string;
  viewCustomStyle: string;

  // Validation Config
  validationConfig: IValidationConfig;
  validationEditorFieldKey?: string;
  // List View Config
  listVisibleFields: string[];
  listFieldOrderMap: { [key: string]: number };
  listFilterMap: { [key: string]: boolean };
  listCustomScript: string;
  listCustomStyle: string;
  // General UI Config
  formLayout: 'single' | 'double';
  themeColor: string;
  // Advanced Features Config
  autocompleteConfig: IAutocompleteConfig;
  autocompleteEditorFieldKey?: string;
  lookupDisplayConfig?: ILookupDisplayConfig;
  cascadeConfig?: ICascadeConfig;
  editCustomActions: ICustomAction[];
  viewCustomActions: ICustomAction[];
  // Property Pane UI State (Not necessarily persisted logic, but UI flags)
  isConfiguringEditActions: boolean;
  isConfiguringViewActions: boolean;
  fieldPermissionConfig: { [fieldKey: string]: string[] }; // Map FieldKey -> Array of Group Names
  permissionEditorFieldKey?: string;

  views: IViewConfig[];
  defaultViewAllowedGroups: string[];
  isConfiguringDefaultViewPerms: boolean;
  enableLogging: boolean;
  showUserAlerts: boolean;
  // ---  PERMISSION OVERRIDES ---
  overrideAdd: boolean;
  overrideEdit: boolean;
  overrideDelete: boolean;
  //FORM SECTION CONFIGURATION
  formSections: IFormSection[];
  sectionLayout: 'none' | 'stacked' | 'tabs' | 'wizard';
  isConfiguringSections: boolean;
  formattingConfig: IFormattingConfig;
  activeFormattingField?: string;
  listPageTitle?: string;
  addPageTitle?: string;
  editPageTitle?: string;
  viewPageTitle?: string;
  addSuccessMessage?: string;
  editSuccessMessage?: string;
  logListTitle: string;
  logRoleName: string;
  isLogInitialized: boolean;
  isLogSetupConfirmed: boolean;
  showHiddenLists: boolean;
  installType: string;
  listGroupByField: string;
  childConfigs: IChildListConfig[];
  isConfiguringChild: boolean;       // [NEW] Tracks if the panel is open
  childConfig_title?: string;        // [NEW] Temp state for title input
  childConfig_list?: string;         // [NEW] Temp state for list dropdown
  childConfig_ref?: string;          // [NEW] Temp state for ref field
  childConfig_mode?: 'row' | 'form';   // [NEW] Temp state for mode
  childConfig_fields?: string;       // [NEW] Temp state for fields input

  repeaterConfig: IRepeaterConfig;

  // --- NOTIFICATION CONFIGURATION ---
  enableNotification: boolean;
  notifRoleName: string;
  // ADD (Create)
  enableNotifAdd: boolean;
  msgNotifAdd: string;
  groupsNotifAdd: any[];
  rulesAdd: INotificationRule[];

  // UPDATE (Edit)
  enableNotifUpdate: boolean;
  msgNotifUpdate: string;
  groupsNotifUpdate: any[];
  rulesUpdate: INotificationRule[];

  // DELETE
  enableNotifDelete: boolean;
  msgNotifDelete: string;
  groupsNotifDelete: any[];
  rulesDelete: INotificationRule[];

  // VIEW (NEW)
  enableNotifView: boolean;
  msgNotifView: string;
  groupsNotifView: any[];
  rulesView: INotificationRule[];
  isConfiguringViews: boolean;
}
// --- HELPER COMPONENT: Multi-Select Dropdown ---
// Used inside the Property Pane for selecting multiple columns
export interface IMultiSelectEditorProps {
  label: string;
  placeholder: string;
  options: { key: string; text: string }[];
  selectedKeys: string[];
  onChanged: (keys: string[]) => void;
}
class MultiSelectEditor extends React.Component<IMultiSelectEditorProps, { selectedKeys: string[] }> {
  constructor(props: IMultiSelectEditorProps) {
    super(props);
    try {
      this.state = { selectedKeys: props.selectedKeys || [] };
    } catch (error: any) {
      void LoggerService.log('MultiSelectEditor-constructor', 'High', 'Config', error.message || JSON.stringify(error));
    }
  }
  public componentWillReceiveProps(nextProps: IMultiSelectEditorProps) {
    try {
      if (nextProps.selectedKeys !== this.props.selectedKeys) {
        this.setState({ selectedKeys: nextProps.selectedKeys || [] });
      }
    } catch (error: any) {
      void LoggerService.log('MultiSelectEditor-willReceiveProps', 'Low', 'Config', error.message || JSON.stringify(error));
    }
  }
  public render() {
    return React.createElement(Dropdown, {
      label: this.props.label,
      multiSelect: true,
      placeHolder: this.props.placeholder,
      options: this.props.options,
      selectedKeys: this.state.selectedKeys,
      // NOTE: Removed 'calloutProps' to fix Layering/Z-Index issues in PropPane
      onChange: (ev: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        try {
          if (!option) return;
          const { selectedKeys } = this.state;
          const newKeys = option.selected
            ? [...selectedKeys, option.key as string]
            : selectedKeys.filter(k => k !== option.key);
          // 1. Update visual state immediately
          this.setState({ selectedKeys: newKeys });
          // 2. Notify parent to save
          this.props.onChanged(newKeys);
        } catch (error: any) {
          void LoggerService.log('MultiSelectEditor-onChange', 'Medium', 'Config', error.message || JSON.stringify(error));
        }
      }
    });
  }
}
// --- HELPER COMPONENT: ColumnMappingEditor ---
// UI for mapping Source Columns -> Target Fields (e.g. for Auto-populate)
export class ColumnMappingEditor extends React.Component<{
  label: string;
  sourceOptions: { key: string; text: string }[];
  targetOptions: { key: string; text: string }[];
  mappings: IColumnMapping[];
  onChanged: (mappings: IColumnMapping[]) => void;
}, { mappings: IColumnMapping[] }> {
  constructor(props: any) {
    super(props);
    try {
      this.state = { mappings: props.mappings || [] };
    } catch (error: any) {
      void LoggerService.log('ColumnMappingEditor-constructor', 'High', 'Config', error.message || JSON.stringify(error));
    }
  }
  public componentWillReceiveProps(nextProps: any) {
    try {
      if (JSON.stringify(nextProps.mappings) !== JSON.stringify(this.props.mappings)) {
        this.setState({ mappings: nextProps.mappings || [] });
      }
    } catch (error: any) {
      void LoggerService.log('ColumnMappingEditor-willReceiveProps', 'Low', 'Config', error.message || JSON.stringify(error));
    }
  }
  private addMapping = () => {
    try {
      // Use .concat instead of spread operator [...] to ensure ES5 compat
      const newMap = this.state.mappings.concat([{ source: '', target: '' }]);
      this.setState({ mappings: newMap }, () => {
        this.props.onChanged(newMap);
      });
    } catch (error: any) {
      void LoggerService.log('ColumnMappingEditor-addMapping', 'Medium', 'Config', error.message || JSON.stringify(error));
    }
  };
  private removeMapping = (index: number) => {
    try {
      const newMap = this.state.mappings.filter((_, i) => i !== index);
      this.setState({ mappings: newMap }, () => {
        this.props.onChanged(newMap);
      });
    } catch (error: any) {
      void LoggerService.log('ColumnMappingEditor-removeMapping', 'Medium', 'Config', error.message || JSON.stringify(error));
    }
  };
  private updateMapping = (index: number, field: 'source' | 'target', val: string) => {
    try {
      const newMap = this.state.mappings.map((item, i) => {
        if (i === index) {
          // We manually create the object to keep ES5 compiler happy
          const updatedItem: IColumnMapping = {
            source: item.source,
            target: item.target
          };
          // Update the specific field
          updatedItem[field] = val || '';
          return updatedItem;
        }
        return item;
      });
      this.setState({ mappings: newMap }, () => {
        this.props.onChanged(newMap);
      });
    } catch (error: any) {
      void LoggerService.log('ColumnMappingEditor-updateMapping', 'Medium', 'Config', error.message || JSON.stringify(error));
    }
  };
  public render() {
    // Renders a list of dropdown pairs (Source -> Target) + Add/Remove buttons
    return React.createElement('div', {
      style: {
        marginTop: 10,
        border: '1px solid #e1dfdd',
        padding: 10,
        background: '#fcfcfc',
        overflowX: 'auto'
      }
    },
      React.createElement('div', { style: { fontWeight: 600, marginBottom: 8 } }, this.props.label),
      // Header Row
      this.state.mappings.length > 0 && React.createElement('div', {
        style: { display: 'flex', gap: '10px', marginBottom: 5, fontSize: '11px', color: '#666', fontWeight: 600, minWidth: '500px' }
      },
        React.createElement('div', { style: { flex: 1 } }, "Source Column"),
        React.createElement('div', { style: { flex: 1 } }, "Target Field"),
        React.createElement('div', { style: { width: '30px' } })
      ),
      // Mapping Rows
      this.state.mappings.map((map, idx) => {
        return React.createElement('div', {
          key: `map-${idx}`,
          style: { display: 'flex', gap: '8px', marginBottom: '10px', alignItems: 'flex-start', minWidth: '500px' }
        },
          // Source Dropdown
          React.createElement('div', { style: { flex: 1 } },
            React.createElement(Dropdown, {
              key: `source-${idx}`,
              options: this.props.sourceOptions,
              selectedKey: map.source || null,
              placeHolder: "Select Source...",
              onChanged: (opt: IDropdownOption) => {
                this.updateMapping(idx, 'source', String(opt.key));
              }
            })
          ),
          // Target Dropdown
          React.createElement('div', { style: { flex: 1 } },
            React.createElement(Dropdown, {
              key: `target-${idx}`,
              options: this.props.targetOptions,
              selectedKey: map.target || null,
              placeHolder: "Select Target...",
              onChanged: (opt: IDropdownOption) => {
                this.updateMapping(idx, 'target', String(opt.key));
              }
            })
          ),
          // Delete Button
          React.createElement('button', {
            onClick: () => this.removeMapping(idx),
            style: {
              marginTop: '4px',
              width: '30px',
              height: '28px',
              border: '1px solid #d13438',
              background: '#fff',
              color: '#d13438',
              fontWeight: 'bold',
              borderRadius: '3px',
              cursor: 'pointer'
            }
          }, "X")
        );
      }),
      // Add Button
      React.createElement('button', {
        onClick: this.addMapping,
        style: {
          marginTop: 8,
          padding: '8px 12px',
          cursor: 'pointer',
          background: '#eff6ff',
          border: '1px solid #0078d4',
          color: '#0078d4',
          borderRadius: '4px',
          width: '100%',
          fontWeight: 600
        }
      }, "+ Add Column Mapping")
    );
  }
}
/**
 * MAIN CLASS: PowerFormWebPart
 * --------------------------------
 * Handles property pane configuration, data loading (fields, lists),
 * and rendering the React container.
 */
export default class PowerFormWebPart extends BaseClientSideWebPart<IPowerFormWebPartProps> {
  private _sp!: SPFI;
  private lists: { key: string; text: string }[] = [];
  private fields: { key: string; text: string; type: string }[] = [];
  private service!: CommonService;
  private showLogViewer: boolean = false;
  // Tracks if the list has >5000 items to switch loading strategies
  private detectedIsLargeList: boolean = false;
  // Variables for Autocomplete/Cascade configuration state
  private sourceLists: { key: string; text: string }[] = [];
  private sourceFields: { key: string; text: string }[] = [];
  private activeAutocompleteField: string | null = null;
  private activeCascadeField: string | null = null;
  private activeLookupConfigField: string | null = null; // Tracks which lookup is being configured
  private activeRepeaterFieldKey: string | null = null;
  private siteGroups: { key: string; text: string }[] = [];

  private async configureNotificationPermissions(): Promise<void> {
    const listName = "BAZnotifications";
    const roleName = this.properties.notifRoleName;

    if (!roleName) {
      void Swal.fire('Error', 'Please enter a Role Name first (e.g. "NotificationSubmitter").', 'error');
      return;
    }

    void Swal.fire({
      title: 'Configuring Permissions...',
      html: 'Checking list and roles...',
      didOpen: () => Swal.showLoading()
    });

    try {
      const web = this._sp.web;

      // 1. Check List
      try {
        await web.lists.getByTitle(listName)();
      } catch (e: any) {
        void Swal.fire('Error', `List '${listName}' not found. Please ensure the feature is activated.`, 'error');
        return;
      }

      // 2. Get Role Definition ID
      const roles = await web.roleDefinitions.filter(`Name eq '${roleName}'`)();
      if (roles.length === 0) {
        void Swal.fire('Error', `Role '${roleName}' not found in this site. Please create the permission level first (Add Items, Edit Items only).`, 'error');
        return;
      }
      const roleDefId = roles[0].Id;

      // 3. Get "Everyone" or "Everyone except external users"
      // This is tricky as names vary. We try standard names.
      let everyoneUser: any = null;
      try {
        // Try resolving commonly used names
        const res = await web.ensureUser("c:0(.s|true"); // "Everyone" claim often looks like this
        everyoneUser = res;
      } catch (e: any) {
        try {
          // Fallback: Try searching for "Everyone" text
          const search = await web.siteUsers.filter("substringof('Everyone', Title)")();
          if (search.length > 0) everyoneUser = search[0];
        } catch (ex) { /* ignore */ }
      }

      if (!everyoneUser) {
        void Swal.fire('warning', 'Could not automatically find "Everyone" group. Please add them manually after this process.', 'warning');
        // We continue to break inheritance at least
      }

      // 4. Break Inheritance (copy=false to remove existing)
      const list = web.lists.getByTitle(listName);
      await list.breakRoleInheritance(false);

      // 5. Add Permission
      if (everyoneUser) {
        await list.roleAssignments.add(everyoneUser.Id, roleDefId);
        void Swal.fire('success', `List permissions reset. 'Everyone' added with role '${roleName}'.`, 'success');
      } else {
        void Swal.fire('success', `List permissions inheritance broken and cleared. Please add the user group manually.`, 'warning');
      }

    } catch (error: any) {
      console.error(error);
      void Swal.fire('Error', 'Failed to configure permissions. Check console.', 'error');
    }
  }

  /**
   * Loads fields for the currently selected SharePoint list.
   * Also detects if the list is "Large" (>5000 items).
   */
  private async loadFieldsForSelectedList(listTitle: string): Promise<void> {
    if (!listTitle) {
      this.fields = [];
      return;
    }
    try {
      // 1. GET LIST METADATA (ItemCount)
      const listMetaUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')?$select=ItemCount`;
      const listMetaRes = await this.context.spHttpClient.get(listMetaUrl, SPHttpClient.configurations.v1);
      const listMeta = await listMetaRes.json();
      const count = listMeta.ItemCount || 0;
      // SharePoint Threshold is 5000. We use >= 5000 to be safe.

      this.detectedIsLargeList = count >= 5000;
      // 2. GET FIELDS
      let url = this.context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getbytitle('" + listTitle + "')/fields?" +
        "$filter=Hidden eq false and (ReadOnlyField eq false or InternalName eq 'Attachments' or InternalName eq 'Created' or InternalName eq 'Modified' or InternalName eq 'Author' or InternalName eq 'Editor')" +
        "&$select=InternalName,Title,TypeAsString,Choices,ReadOnlyField,Required";
      let res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      let data: any = await res.json();
      let arr = data.value || [];
      this.fields = arr.filter(function (f: any) { return f && f.InternalName; })
        .map(function (f: any) {
          return { key: f.InternalName, text: f.Title || f.InternalName, type: f.TypeAsString, choices: f.Choices || [], Required: f.Required };
        });
    } catch (e: any) {
      void LoggerService.log('PowerFormWebPart-loadFieldsForSelectedList', 'High', 'Config', e.message || JSON.stringify(e));
    }
  }
  private async loadSharePointGroups(): Promise<void> {
    try {
      // Fetch ID and Title of site groups
      const groups = await this._sp.web.siteGroups.select('Id', 'Title')();
      this.siteGroups = groups.map((g: any) => ({
        key: String(g.Id),
        text: g.Title
      })).sort((a, b) => a.text.localeCompare(b.text));
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-loadSharePointGroups', 'Medium', 'Config', error.message);
    }
  }
  public async onInit(): Promise<void> {
    try {
      const style = document.createElement('style');
      style.type = 'text/css';
      style.innerHTML = `
    /* GENERIC SELECTOR: Targets any class starting with 'propertyPanePageDescription'  */
    div[class^="propertyPanePageDescription_"], 
    div[class*=" propertyPanePageDescription_"] {
      font-weight: bold !important;
      font-size: x-large !important;
      color: #323130 !important;
      margin-bottom: 15px !important;
      display: block !important;
    }
  `;
      document.head.appendChild(style);
      this.service = new CommonService(this.context.pageContext.site.absoluteUrl, this.context);
      // 1. Ensure Log List Exists (Background Task)
      //LoggerService.ensureLogList(this.context);
      // 2. Load Lists
      await this.loadLists();
      await this.loadSharePointGroups();
      // 3. Load Fields if list is selected
      if (this.properties.selectedList) {
        await this.loadFieldsForSelectedList(this.properties.selectedList);
      }
    } catch (error: any) {
      // Cannot log to SPFx list if init fails, but we try anyway
      void LoggerService.log('PowerFormWebPart-onInit', 'High', 'Config', error.message || JSON.stringify(error));
    }
    return super.onInit();
  }
  /** Fetch all non-hidden lists from the current site */
  private async loadLists(): Promise<void> {
    try {
      // DEFAULT: Visible Custom Lists only
      let filterQuery = "Hidden eq false and BaseTemplate eq 100";

      // IF TOGGLE ON: Show ALL Custom Lists (Visible + Hidden)
      if (this.properties.showHiddenLists) {
        filterQuery = "BaseTemplate eq 100";
      }

      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=${filterQuery}&$select=Title,Id`;
      const res = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

      let data: any = await res.json();
      let arr = data.value || [];
      this.lists = arr.map(function (l: any) {
        return { key: l.Title || String(l.Id), text: l.Title || String(l.Id) };
      });
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-loadLists', 'High', 'Config', error.message || JSON.stringify(error));
    }
  }
  /** * Property Pane Configuration Start
   * Injects custom CSS to widen the Property Pane panel for better UX.
   */
  protected async onPropertyPaneConfigurationStart(): Promise<void> {

    try {
      // FORCE WIDTH: Target all possible Property Pane containers using Wildcard Selectors
      const css = `
      /* 1. Page Header (General Settings) */
      
      .spPropertyPaneContainer .ms-PropertyPaneGroup-header,
.spPropertyPaneContainer .ms-PropertyPaneGroup-title,
button[class*="header"] span[class*="headerText"] {
  font-weight: 800 !important;
  color: #000 !important;
}
      .ms-PropertyPanePage-header {
    font-weight: 800 !important;
  }
  /* Bold all Accordion Section Headers (Add, Edit, View, List, etc.) */
  .ms-PropertyPaneGroup-header, 
  .ms-PropertyPaneGroup-title,
  button[class*="header"] span {
    font-weight: 800 !important;
    color: #000 !important;
  }
  /* Bold all Field Labels (List, Theme Color, etc.) */
  .ms-Label {
    font-weight: 800 !important;
  }
      .spPropertyPaneContainer .ms-PropertyPaneGroup-header,
.spPropertyPaneContainer .ms-PropertyPaneGroup-title,
.spPropertyPaneContainer .ms-PropertyPanePage-header,
.spPropertyPaneContainer .ms-Label,
div[class*="spPropertyPaneContainer"] span[class*="header"],
div[class*="spPropertyPaneContainer"] button[class*="header"] {
  font-weight: 800 !important;
  color: #000 !important;
}
        .spPropertyPaneContainer .ms-PropertyPanePage-header,
        div[class*="spPropertyPaneContainer"] .ms-PropertyPanePage-header {
          font-weight: 800 !important;
          font-size: 17px !important;
          color: #000 !important;
        }
        /* 2. Group Headers (Add/Edit/View/List Configuration, Maintenance etc.) */
        .spPropertyPaneContainer .ms-PropertyPaneGroup-header,
        .spPropertyPaneContainer .ms-PropertyPaneGroup-title,
        div[class*="spPropertyPaneContainer"] .ms-PropertyPaneGroup-header {
          font-weight: 700 !important;
          color: #333 !important;
        }
        /* 3. Field Labels (List, Theme Color, Custom JavaScript) */
        .spPropertyPaneContainer .ms-Label,
        div[class*="spPropertyPaneContainer"] .ms-Label {
          font-weight: 700 !important;
          color: #000 !important;
        }
        /* 4.  the Panel alignment (Your existing logic) */
        .spPropertyPaneContainer .ms-Panel-main,
        div[class*="spPropertyPaneContainer"] .ms-Panel-main {
          width: 100% !important;
          right: 0 !important; 
        }
      
        /* 5. Ensure Content Area fills the space */
        .spPropertyPaneContainer .ms-Panel-contentInner,
        div[class*="spPropertyPaneContainer"] .ms-Panel-contentInner {
          width: 100% !important;
        }
        
        /* 6. Allow horizontal scrolling generally inside the panel content */
        .spPropertyPaneContainer .ms-Panel-scrollableContent,
        div[class*="spPropertyPaneContainer"] .ms-Panel-scrollableContent {
          overflow-x: hidden; /* Hide main scrollbar, let specific tables scroll */
          width: 100% !important;
        }
      `;
      //Ensure CSS is injected
      if (!document.getElementById('custom-pp-width-FIXED')) {
        const head = document.getElementsByTagName('head')[0];
        const style = document.createElement('style');
        style.setAttribute('id', 'custom-pp-width-FIXED');
        style.type = 'text/css';
        style.appendChild(document.createTextNode(css));
        head.appendChild(style);
      }
      if (this.properties.selectedList) {
        await this.loadFieldsForSelectedList(this.properties.selectedList);
        this.context.propertyPane.refresh();
      }
      if (this.siteGroups.length === 0) {
        void this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/sitegroups?$select=Id,Title`,
          SPHttpClient.configurations.v1
        )
          .then(response => response.json())
          .then(data => {
            if (data && data.value) {
              // Map groups to Dropdown Options
              this.siteGroups = data.value.map((g: any) => ({ key: g.Id.toString(), text: g.Title }));

              // Add a blank/default option at the top
              this.siteGroups.unshift({ key: '', text: '-- Everyone (No Group) --' });

              this.context.propertyPane.refresh();
            }
          })
          .catch(err => console.error("Error fetching groups:", err));
      }
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-onPropertyPaneConfigurationStart', 'Medium', 'Config', error.message || JSON.stringify(error));
    }
  }
  /** Handle changes in Property Pane fields (e.g., list selection) */
  protected async onPropertyPaneFieldChanged(path: string, oldVal: any, newVal: any): Promise<void> {
    try {
      // 1. Handle Main List Change
      if (path === 'showHiddenLists') {
        // Clear current selection if it might disappear
        this.properties.selectedList = '';
        // Reload lists with new filter
        void this.loadLists();
      }
      if (path === 'selectedList' && oldVal !== newVal) {
        await this.loadFieldsForSelectedList(newVal);
        // Reset visible fields to default (All)
        const allFields = this.fields.map(f => f.key);
        this.properties.addVisibleFields = [...allFields];
        this.properties.editVisibleFields = [...allFields];
        this.properties.viewVisibleFields = [...allFields];
        this.properties.listVisibleFields = allFields;
        // Clear configurations associated with the previous list
        this.properties.addFieldOrder = {};
        this.properties.addReadOnlyFields = [];
        this.properties.addCustomScript = '';
        this.properties.addCustomStyle = '';
        this.properties.editFieldOrder = {};
        this.properties.editReadOnlyFields = [];
        this.properties.editCustomScript = '';
        this.properties.editCustomStyle = '';
        this.properties.viewFieldOrder = {};
        this.properties.viewCustomScript = '';
        this.properties.viewCustomStyle = '';
        this.properties.listFieldOrderMap = {};
        this.properties.listFilterMap = {};
        this.properties.validationEditorFieldKey = undefined;
        this.context.propertyPane.refresh();
        this.render();
      }
      // 2. Handle Autocomplete Source List Change (Refresh Source Fields dropdown)
      if (path.indexOf('autocompleteConfig') > -1 && path.indexOf('sourceList') > -1) {
        // newVal is the List ID/Title selected
        if (newVal) {
          // Show loading indicator (optional optimization) or just await
          await this.loadSourceFields(newVal);
          // Refresh to show the newly loaded columns in the "Main Display Column" dropdown
          this.context.propertyPane.refresh();
        }
      }
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-onPropertyPaneFieldChanged', 'Medium', 'Config', error.message || JSON.stringify(error));
    }
    super.onPropertyPaneFieldChanged(path, oldVal, newVal);
  }
  protected render(): void {
    try {
      // 1. UPDATE LOGGER STATE
      LoggerService.init(
        this.context,
        this.properties.selectedList,
        this.properties.logListTitle,
        this.properties.enableLogging,
        this.properties.showUserAlerts
      );
      // 2. WELCOME SCREEN (If no list is configured)
      if (!this.properties.selectedList) {
        const isSiteAdmin = this.context.pageContext.legacyPageContext['isSiteAdmin'];
        if (!isSiteAdmin) {
          const nonAdminElement = React.createElement('div', {
            style: {
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              minHeight: '300px',
              padding: '40px',
              textAlign: 'center',
              backgroundColor: '#fff4f4', // Light red background
              borderRadius: '8px',
              border: '1px solid #ffcccc',

              boxShadow: '0 2px 10px rgba(0,0,0,0.05)',
              margin: '20px'
            }
          },
            React.createElement('div', { style: { fontSize: '48px', marginBottom: '15px' } }, '🔒'),
            React.createElement('h2', { style: { margin: '0 0 10px 0', color: '#d13438', fontWeight: 600 } }, 'Setup Required'),
            React.createElement('p', { style: { margin: '0 0 10px 0', color: '#333', maxWidth: '500px', lineHeight: '1.6', fontSize: '15px' } },
              'This web part requires initial configuration and list creation.'
            ),
            React.createElement('p', { style: { margin: '0', color: '#666', maxWidth: '500px', fontSize: '14px', fontStyle: 'italic' } },
              'Please contact your Site Administrator to add and configure this web part.'
            )
          );
          ReactDom.render(nonAdminElement, this.domElement);
          return;
        }
        const welcomeElement = React.createElement('div', {
          style: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            minHeight: '300px',
            padding: '40px',
            textAlign: 'center',
            backgroundColor: '#ffffff',
            borderRadius: '8px',

            boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
            margin: '20px'
          }
        },
          React.createElement('div', { style: { fontSize: '64px', marginBottom: '20px' } }, '📝'),
          React.createElement('h2', { style: { margin: '0 0 10px 0', color: '#333', fontWeight: 600 } }, 'Welcome to BAZ Dynamic Form'),
          React.createElement('p', { style: { margin: '0 0 20px 0', color: '#666', maxWidth: '500px', lineHeight: '1.6', fontSize: '15px' } },
            'The form is not configured yet. Please configure it in order to use and enjoy the features of BAZ Dynamic Form.'
          ),
          React.createElement('div', { style: { fontSize: '13px', color: '#888', marginTop: '10px', padding: '10px', background: '#f5f5f5', borderRadius: '4px' } },
            'If there is any query please contact ',
            React.createElement('a', { href: 'mailto:admin@baztechnologies.ae', style: { color: '#0078d4', textDecoration: 'none', fontWeight: 600 } }, 'admin@baztechnologies.ae')
          ),
          React.createElement('button', {
            style: {
              marginTop: '30px',
              padding: '12px 32px',
              backgroundColor: '#0078d4',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontWeight: 600,
              fontSize: '14px',
              boxShadow: '0 4px 6px rgba(0,120,212,0.3)',
              transition: 'background 0.2s'
            },
            onClick: () => { this.context.propertyPane.open(); }
          }, 'Configure Now')
        );
        ReactDom.render(welcomeElement, this.domElement);
        return;
      }
      let activeThemeColor = this.properties.themeColor || '#3b82f6';
      if (activeThemeColor === 'siteTheme') {
        // Access the site's primary theme color (themePrimary)
        activeThemeColor = (window as any).__themeState__ &&
          (window as any).__themeState__.theme &&
          (window as any).__themeState__.theme.themePrimary
          ? (window as any).__themeState__.theme.themePrimary
          : '#0078d4'; // Fallback to standard SP Blue
      }
      // 3. CREATE FORM ELEMENT
      const formElement = React.createElement<any>(PowerForm, {
        selectedList: this.properties.selectedList,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        themeColor: this.properties.themeColor || '#3b82f6',
        // Ensure this property exists in your class, otherwise remove it or set default false
        isLargeList: this.properties.isLargeList || (this as any).detectedIsLargeList || false,
        formLayout: this.properties.formLayout,
        // Config Props
        addVisibleFields: this.properties.addVisibleFields || [],
        addFieldOrder: this.properties.addFieldOrder || {},
        addReadOnlyFields: this.properties.addReadOnlyFields || [],
        addCustomScript: this.properties.addCustomScript || '',
        addCustomStyle: this.properties.addCustomStyle || '',
        editVisibleFields: this.properties.editVisibleFields || [],
        editFieldOrder: this.properties.editFieldOrder || {},
        editReadOnlyFields: this.properties.editReadOnlyFields || [],
        editCustomScript: this.properties.editCustomScript || '',
        editCustomStyle: this.properties.editCustomStyle || '',
        viewVisibleFields: this.properties.viewVisibleFields || [],
        viewFieldOrder: this.properties.viewFieldOrder || {},
        viewCustomScript: this.properties.viewCustomScript || '',
        viewCustomStyle: this.properties.viewCustomStyle || '',
        listVisibleFields: this.properties.listVisibleFields || [],
        listFieldOrderMap: this.properties.listFieldOrderMap || {},
        listFilterMap: this.properties.listFilterMap || {},
        listCustomScript: this.properties.listCustomScript || '',
        listCustomStyle: this.properties.listCustomStyle || '',
        // Advanced Configs
        autocompleteConfig: this.properties.autocompleteConfig,
        cascadeConfig: this.properties.cascadeConfig,
        lookupDisplayConfig: this.properties.lookupDisplayConfig,
        defaultViewAllowedGroups: this.properties.defaultViewAllowedGroups || [],
        views: this.properties.views || [],
        validationConfig: this.properties.validationConfig || {},
        editCustomActions: this.properties.editCustomActions || [],
        viewCustomActions: this.properties.viewCustomActions || [],
        fieldPermissionConfig: this.properties.fieldPermissionConfig || {},
        listGroupByField: this.properties.listGroupByField,
        childConfigs: this.properties.childConfigs,
        repeaterConfig: this.properties.repeaterConfig,
        service: (this as any).service, // Cast to any if service is protected
        // --- NEW: PASS OVERRIDES ---
        overrideAdd: this.properties.overrideAdd,
        overrideEdit: this.properties.overrideEdit,
        overrideDelete: this.properties.overrideDelete,
        formSections: this.properties.formSections || [],
        sectionLayout: this.properties.sectionLayout || 'stacked',
        formattingConfig: this.properties.formattingConfig || {},
        listPageTitle: this.properties.listPageTitle,
        addPageTitle: this.properties.addPageTitle,
        editPageTitle: this.properties.editPageTitle,
        viewPageTitle: this.properties.viewPageTitle,
        addSuccessMessage: this.properties.addSuccessMessage,
        editSuccessMessage: this.properties.editSuccessMessage,
        enableNotification: this.properties.enableNotification,



        // --- ADD ---
        enableNotifAdd: this.properties.enableNotifAdd, msgNotifAdd: this.properties.msgNotifAdd, groupsNotifAdd: this.properties.groupsNotifAdd,
        rulesAdd: this.properties.rulesAdd,// UI Toggle
        enableNotifUpdate: this.properties.enableNotifUpdate, msgNotifUpdate: this.properties.msgNotifUpdate, groupsNotifUpdate: this.properties.groupsNotifUpdate,
        rulesUpdate: this.properties.rulesUpdate, // UI Toggle

        // --- DELETE ---
        enableNotifDelete: this.properties.enableNotifDelete, msgNotifDelete: this.properties.msgNotifDelete, groupsNotifDelete: this.properties.groupsNotifDelete,
        rulesDelete: this.properties.rulesDelete, // UI Toggle


        enableNotifView: this.properties.enableNotifView, msgNotifView: this.properties.msgNotifView, groupsNotifView: this.properties.groupsNotifView,
        rulesView: this.properties.rulesView, // UI Toggle
        context: this.context,
        initialMode: 'list'
      });
      // 4. CREATE LOG VIEWER ELEMENT
      const logViewerElement = React.createElement(LogViewer, {
        isOpen: (this as any).showLogViewer === true,
        context: this.context,
        currentListTitle: this.properties.selectedList,
        onDismiss: () => {
          (this as any).showLogViewer = false;
          this.render();
        }
      });
      // 5. RENDER BOTH (Wrapped in DIV to fix React.Fragment error)
      ReactDom.render(
        React.createElement('div', null, formElement, logViewerElement),
        this.domElement
      );
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-render', 'High', 'Config', error.message || JSON.stringify(error));
    }
  }
  // --- PROPERTY PANE HELPER METHODS ---
  private getConfigurationActionGroup_hold() {
    return {
      groupName: 'Panel Actions',
      groupFields: [
        PropertyPaneButton('btnSaveConfig', {
          text: 'Save Configuration',
          buttonType: PropertyPaneButtonType.Primary,
          icon: 'Save',
          onClick: () => {
            void Swal.fire({ icon: 'success', title: 'success', text: "Configuration saved successfully!" });
            (this.context.propertyPane as any).close();
          }
        }) as any,
        PropertyPaneButton('btnCancelConfig', {
          text: 'Cancel & Reload',
          buttonType: PropertyPaneButtonType.Normal,
          icon: 'Cancel',
          onClick: () => {
            if (confirm('Are you sure? This will reload the page and revert unsaved UI changes.')) {
              window.location.reload(); // Refresh the page
            }
          }
        }) as any
      ]
    };
  }

  private getConfigurationActionGroup() {
    return {
      groupName: 'Panel Actions',
      groupFields: [
        PropertyPaneButton('btnSaveConfig', {
          text: 'Save Configuration',
          buttonType: PropertyPaneButtonType.Primary,
          icon: 'Save',
          onClick: () => {
            // 1. FORCE COMMIT: Blur any active input
            if (document.activeElement instanceof HTMLElement) {
              document.activeElement.blur();
            }

            // 2. SAFETY CHECK: Access variables correctly based on your file structure

            // A. Properties from this.properties (Persisted)
            // We use bracket notation ['key'] for properties not in the Interface to avoid TS errors
            const isConfiguringSections = this.properties.isConfiguringSections;
            const validationEditorFieldKey = this.properties.validationEditorFieldKey;
            const permissionEditorFieldKey = this.properties.permissionEditorFieldKey;
            const isConfiguringViews = this.properties['isConfiguringViews'];
            const activeFormattingField = this.properties.activeFormattingField;
            const isConfiguringChild = this.properties.isConfiguringChild;

            // B. Class Members (UI State, not persisted)
            // This was the cause of the error - it lives on 'this', not 'this.properties'
            const activeRepeaterFieldKey = this.activeRepeaterFieldKey;

            if (validationEditorFieldKey ||
              permissionEditorFieldKey ||
              isConfiguringSections ||
              activeRepeaterFieldKey ||
              isConfiguringViews ||
              activeFormattingField ||
              isConfiguringChild
            ) {

              void Swal.fire({
                icon: 'warning',
                title: 'Unsaved Changes',
                text: 'You have a configuration panel open (Validation, Sections, Views, etc.). Please save or cancel that panel before saving the main configuration.',
                confirmButtonText: 'OK, I will check'
              });
              return; // STOP
            }

            // 3. SUCCESS & CLOSE
            setTimeout(() => {
              void Swal.fire({
                icon: 'success',
                title: 'Saved',
                text: "Configuration has been applied successfully!",
                timer: 1500,
                showConfirmButton: false
              });
              (this.context.propertyPane as any).close();
            }, 200);
          }
        }) as any,

        PropertyPaneButton('btnCancelConfig', {
          text: 'Cancel & Reload',
          buttonType: PropertyPaneButtonType.Normal,
          icon: 'Cancel',
          onClick: () => {
            void Swal.fire({
              title: 'Discard changes?',
              text: "Any unsaved changes will be lost. The page will reload.",
              icon: 'warning',
              showCancelButton: true,
              confirmButtonColor: '#d33',
              cancelButtonColor: '#3085d6',
              confirmButtonText: 'Yes, discard changes'
            }).then((result: any) => { // Type 'any' fixes the TS error
              if (result.isConfirmed) {
                window.location.reload();
              }
            });
          }
        }) as any
      ]
    };
  }
  /** Renders the Custom Action configuration (Edit/View buttons) */
  private getActionConfigControl(type: 'edit' | 'view'): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    try {
      const isEditing = type === 'edit' ? this.properties.isConfiguringEditActions : this.properties.isConfiguringViewActions;
      const currentActions = type === 'edit' ? (this.properties.editCustomActions || []) : (this.properties.viewCustomActions || []);
      if (isEditing) {
        // Show Editor
        return PropertyPaneCustomField({
          key: `actionEditor_${type} `,
          onRender: (elem) => {
            const editor = React.createElement(CustomActionEditor, {
              label: `Configure ${type === 'edit' ? 'Edit' : 'View'} Actions`,
              actions: currentActions,
              onSave: (newActions: ICustomAction[]) => {
                if (type === 'edit') {
                  this.properties.editCustomActions = newActions;
                  this.properties.isConfiguringEditActions = false;
                } else {
                  this.properties.viewCustomActions = newActions;
                  this.properties.isConfiguringViewActions = false;
                }
                this.context.propertyPane.refresh();
              },
              onCancel: () => {
                if (type === 'edit') this.properties.isConfiguringEditActions = false;
                else this.properties.isConfiguringViewActions = false;
                this.context.propertyPane.refresh();
              }
            });
            ReactDom.render(editor, elem);
          }
        });
      }
      // 2. Otherwise, Show "Configure Buttons" Button
      return PropertyPaneCustomField({
        key: `btnLaunch_${type} `,
        onRender: (elem) => {
          const btn = React.createElement('button', {
            onClick: () => {
              if (type === 'edit') this.properties.isConfiguringEditActions = true;
              else this.properties.isConfiguringViewActions = true;
              this.context.propertyPane.refresh();
            },
            style: {
              padding: '8px 16px', cursor: 'pointer', backgroundColor: '#fff',
              border: '1px solid #0078d4', color: '#0078d4', borderRadius: '4px', width: '100%'
            }
          }, `Configure ${type === 'edit' ? 'Edit' : 'View'} Buttons(${currentActions.length})`);
          ReactDom.render(btn, elem);
        }
      });
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-getActionConfigControl', 'Medium', 'Config', error.message || JSON.stringify(error));
      return PropertyPaneLabel('err', { text: 'Error loading Action config.' }) as any;
    }
  }
  // Helper to map property names based on the current mode (Add/Edit/View/List)
  private getPropNames(type: 'add' | 'edit' | 'view' | 'list') {
    switch (type) {
      case 'add': return { vis: 'addVisibleFields', ord: 'addFieldOrder', ro: 'addReadOnlyFields' };
      case 'edit': return { vis: 'editVisibleFields', ord: 'editFieldOrder', ro: 'editReadOnlyFields' };
      case 'view': return { vis: 'viewVisibleFields', ord: 'viewFieldOrder', ro: null };
      case 'list': return { vis: 'listVisibleFields', ord: 'listFieldOrderMap', ro: null };
    }
  }
  /** Renders a "Select All" / "Unselect All" button for field configuration */
  private getSelectAllControl(type: 'add' | 'edit' | 'view' | 'list'): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    return PropertyPaneCustomField({
      key: 'selectAll_' + type,
      onRender: (elem) => {
        try {
          const props = this.getPropNames(type);
          const currentSelected: string[] = (this.properties as any)[props.vis] || [];
          const allKeys = this.fields.filter(f => f.key !== 'ContentType').map(f => f.key);
          const isAllSelected = allKeys.length > 0 && allKeys.every(k => currentSelected.indexOf(k) > -1);
          const onToggleAll = () => {
            if (isAllSelected) {
              // Unselect All
              (this.properties as any)[props.vis] = [];
            } else {
              // Select All
              (this.properties as any)[props.vis] = [...allKeys];
            }
            this.context.propertyPane.refresh();
          };
          const btnStyle: React.CSSProperties = {
            cursor: 'pointer',
            background: isAllSelected ? '#fee2e2' : '#dbeafe',
            color: isAllSelected ? '#b91c1c' : '#1e40af',
            border: '1px solid ' + (isAllSelected ? '#f87171' : '#60a5fa'),
            padding: '6px 12px',
            borderRadius: '4px',
            fontSize: '12px',
            marginBottom: '15px',
            width: '100%',
            fontWeight: 600 as any,
            textAlign: 'center',
            boxSizing: 'border-box'
          };
          const btn = React.createElement('div', {
            style: btnStyle,
            onClick: onToggleAll
          }, isAllSelected ? 'Unselect All Fields' : 'Select All Fields');
          ReactDom.render(btn, elem);
        } catch (error: any) {
          void LoggerService.log('PowerFormWebPart-getSelectAllControl', 'Medium', 'Config', error.message || JSON.stringify(error));
        }
      }
    });
  }
  /**
   * Fetches fields from a target "Source List" (used for Autocomplete/Lookups).
   * Handles both List Title and List GUID.
   */
  private async loadSourceFields(listIdOrTitle: string) {
    try {
      // 1. Clean the ID (Remove curly braces if present)
      const cleanId = listIdOrTitle.replace(/^{|}$/g, '');
      // 2. Check if it looks like a GUID (8-4-4-4-12 hex pattern)
      const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(cleanId);
      // 3. Construct Endpoint (No spaces!)
      const endpoint = isGuid
        ? `lists(guid'${cleanId}')`
        : `lists/getbytitle('${listIdOrTitle}')`;
      // 4. Fetch Fields (Removed 'TypeAsString eq Text' so you can see Number/Date columns too)
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/${endpoint}/fields?$filter=Hidden eq false and ReadOnlyField eq false`;
      const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await res.json();
      // 5. Map to Dropdown Options
      this.sourceFields = (data.value || [])
        .filter((f: any) => {
          const name = f.InternalName;
          // Exclude specific system columns
          return name !== 'ContentType' &&
            name !== 'Attachments' &&
            name !== 'Edit' &&
            name !== 'DocIcon' &&
            name !== 'ItemChildCount' &&
            name !== 'FolderChildCount';
        })
        .map((f: any) => ({
          key: f.InternalName,
          text: `${f.Title} (${f.InternalName})`
        }));
      // Sort alphabetically
      this.sourceFields.sort((a, b) => a.text.localeCompare(b.text));
    } catch (e: any) {
      void LoggerService.log('PowerFormWebPart-loadSourceFields', 'Medium', 'Config', e.message || JSON.stringify(e));
      this.sourceFields = [];
    }
  }
  /**
   * Constructs the controls for a single field row in the Property Pane.
   * Includes visibility toggle, order buttons, and configuration icons (Validation, Permissions, etc.)
   */
  private getFieldControl(
    field: { key: string; text: string; type: string; Required?: boolean },
    maxOrder: number,
    type: 'add' | 'edit' | 'view' | 'list'
  ): IPropertyPaneField<IPropertyPaneCustomFieldProps>[] {
    try {
      const props = this.getPropNames(type);

      // State flags for current field
      const activeKey = this.properties.validationEditorFieldKey;
      const isEditingValidation = (type === 'add' || type === 'edit') && activeKey === field.key;
      const isConfiguringAutocomplete = this.activeAutocompleteField === field.key;
      const isConfiguringLookup = this.activeLookupConfigField === field.key;
      const isConfiguringCascade = this.activeCascadeField === field.key;
      const activePermKey = this.properties.permissionEditorFieldKey;
      const isEditingPerms = (type === 'add' || type === 'edit') && activePermKey === field.key;

      let formattingEditor = null;
      const isRequired = field.Required === true;
      const isLocked = type === 'add' && isRequired;
      const displayLabel = isLocked ? `${field.text} * (Required)` : field.text;

      // 1. MAIN FIELD ROW (Rendered via CustomField)
      const row = PropertyPaneCustomField({
        key: 'row_' + type + '_' + field.key,
        onRender: (elem: HTMLElement) => {
          try {
            const visArr = (this.properties as any)[props.vis] || [];
            const checked = isLocked ? true : (visArr.indexOf(field.key) !== -1);
            const ordMap = (this.properties as any)[props.ord] || {};
            const order = ordMap[field.key] != null ? ordMap[field.key] : 1;

            // Toggle Visibility Handler - Using Arrow Function to fix 'this' context
            const onToggle = (k: string, t: string, chk: boolean) => {
              const s = new Set((this.properties as any)[props.vis] || []);
              if (chk) {
                s.add(k);
              } else {
                s.delete(k);
              }
              (this.properties as any)[props.vis] = setToArray(s);
              this.context.propertyPane.refresh();
            };

            // Change Order Handler - Using Arrow Function to fix shadowing
            const onOrder = (k: string, newPos: number) => {
              const visibleKeys = ((this.properties as any)[props.vis] || []).slice();
              const currentMap = (this.properties as any)[props.ord] || {};

              visibleKeys.sort((a: string, b: string) => {
                const valA = currentMap[a] !== undefined ? currentMap[a] : 999;
                const valB = currentMap[b] !== undefined ? currentMap[b] : 999;
                return valA - valB;
              });

              const idx = visibleKeys.indexOf(k);
              if (idx > -1) { visibleKeys.splice(idx, 1); }
              const insertIdx = Math.max(0, Math.min(newPos - 1, visibleKeys.length));
              visibleKeys.splice(insertIdx, 0, k);

              const newMap: { [key: string]: number } = {};
              for (let i = 0; i < visibleKeys.length; i++) {
                newMap[visibleKeys[i]] = i + 1;
              }
              (this.properties as any)[props.ord] = newMap;
              (this.properties as any)[props.vis] = visibleKeys;
              this.context.propertyPane.refresh();
            };

            const children: any[] = [
              React.createElement(FieldOrderRenderer as any, {
                fieldKey: field.key,
                fieldTitle: displayLabel,
                selected: checked,
                order: order,
                max: maxOrder,
                disabled: isLocked,
                onToggle: (k: string, chk: boolean) => {
                  if (isLocked) return;
                  onToggle(k, field.text, chk);
                },
                onOrderChange: onOrder
              } as any)
            ];

            const iconBtnStyle = (isActive: boolean, activeColor: string = '#0078d4') => ({
              marginLeft: 6,
              fontSize: '14px',
              cursor: 'pointer',
              background: isActive ? '#eff6ff' : 'transparent',
              border: isActive ? `1px solid ${activeColor}` : '1px solid transparent',
              color: isActive ? activeColor : '#666',
              padding: '4px',
              borderRadius: '4px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: '28px',
              height: '28px'
            });

            // 1. Read Only Toggle (Fixed TS2538 Index Type error)
            const roProp = props.ro;
            if (roProp) {
              const roArr = (this.properties as any)[roProp] || [];
              const isRo = roArr.indexOf(field.key) !== -1;

              const onToggleRo = (e: any) => {
                if (e) e.preventDefault();
                if (isLocked) return;
                const s = new Set((this.properties as any)[roProp] || []);
                if (s.has(field.key)) s.delete(field.key);
                else s.add(field.key);
                (this.properties as any)[roProp] = setToArray(s);
                this.context.propertyPane.refresh();
              };

              const lockedStyle: any = iconBtnStyle(isRo, '#d13438');
              if (isLocked) {
                lockedStyle.opacity = '0.3';
                lockedStyle.cursor = 'not-allowed';
                lockedStyle.borderColor = '#eaeaea';
              }
              children.push(
                React.createElement('button', {
                  style: lockedStyle,
                  onClick: onToggleRo,
                  disabled: isLocked,
                  title: isLocked ? "Required fields cannot be Read-Only" : (isRo ? "Read Only: ON" : "Read Only: OFF")
                }, isRo ? React.createElement(Icons.Lock) : React.createElement(Icons.Unlock))
              );
            }

            // 2. Validation Button
            if (type === 'add' || type === 'edit') {
              const onEditVal = () => {
                this.activeLookupConfigField = null;
                this.activeAutocompleteField = null;
                this.activeCascadeField = null;
                this.properties.permissionEditorFieldKey = undefined;
                this.properties.validationEditorFieldKey = field.key;
                this.context.propertyPane.open();
                this.context.propertyPane.refresh();
              };
              const hasRules = this.properties.validationConfig &&
                this.properties.validationConfig[field.key] &&
                this.properties.validationConfig[field.key].length > 0;
              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(hasRules || isEditingValidation),
                  onClick: onEditVal,
                  title: "Validation Rules"
                }, React.createElement(Icons.Validation))
              );
            }

            // 3. Permissions Button
            if (type === 'add' || type === 'edit') {
              const onEditPerms = () => {
                this.activeLookupConfigField = null;
                this.activeAutocompleteField = null;
                this.activeCascadeField = null;
                this.properties.validationEditorFieldKey = undefined;
                this.properties.permissionEditorFieldKey = field.key;
                this.context.propertyPane.refresh();
              };
              const hasPerms = this.properties.fieldPermissionConfig &&
                this.properties.fieldPermissionConfig[field.key] &&
                this.properties.fieldPermissionConfig[field.key].length > 0;
              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(hasPerms || isEditingPerms, '#107c10'),
                  onClick: onEditPerms,
                  title: "Field Level Permissions"
                }, React.createElement(Icons.Permissions))
              );
            }

            // 4. Lookup Config Button
            if ((type === 'add' || type === 'edit') && (field.type === 'Lookup' || field.type === 'LookupMulti')) {
              const hasLookupCfg = this.properties.lookupDisplayConfig &&
                this.properties.lookupDisplayConfig[field.key] &&
                this.properties.lookupDisplayConfig[field.key].additionalFields.length > 0;

              const onLookupConfigClick = async () => {
                if (this.activeCascadeField === field.key) {
                  this.activeCascadeField = null;
                  this.sourceFields = [];
                } else {
                  this.activeCascadeField = null;
                  this.activeAutocompleteField = null;
                  this.properties.validationEditorFieldKey = undefined;
                  this.properties.permissionEditorFieldKey = undefined;
                  this.activeLookupConfigField = field.key;
                  try {
                    const fDefUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/fields/getByInternalNameOrTitle('${field.key}')`;
                    const fRes = await this.context.spHttpClient.get(fDefUrl, SPHttpClient.configurations.v1);
                    const fData = await fRes.json();
                    if (fData.LookupList) {
                      await this.loadSourceFields(fData.LookupList);
                    }
                  } catch (e: any) {
                    void LoggerService.log('PowerFormWebPart-LookupConfigClick', 'Medium', 'Config', e.message || JSON.stringify(e));
                  }
                }
                this.context.propertyPane.refresh();
              };
              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(isConfiguringLookup || !!hasLookupCfg, '#8764b8'),
                  onClick: onLookupConfigClick,
                  title: "Configure Display Columns"
                }, React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("path", { d: "M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" })))
              );
            }

            // 5. Cascade Button
            if ((type === 'add' || type === 'edit') && (field.type === 'Lookup' || field.type === 'LookupMulti')) {
              const onCascadeClick = async () => {
                if (this.activeCascadeField === field.key) {
                  this.activeCascadeField = null;
                  this.sourceFields = [];
                } else {
                  this.activeLookupConfigField = null;
                  this.activeAutocompleteField = null;
                  this.properties.validationEditorFieldKey = undefined;
                  this.properties.permissionEditorFieldKey = undefined;
                  this.activeCascadeField = field.key;
                  try {
                    const fDefUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/fields/getByInternalNameOrTitle('${field.key}')`;
                    const fRes = await this.context.spHttpClient.get(fDefUrl, SPHttpClient.configurations.v1);
                    const fData = await fRes.json();
                    if (fData.LookupList) {
                      await this.loadSourceFields(fData.LookupList);
                    } else {
                      void LoggerService.log('PowerFormWebPart-onCascadeClick', 'Medium', 'Config', 'No LookupList found for field ' + field.key);
                    }
                  } catch (e: any) {
                    void LoggerService.log('PowerFormWebPart-onCascadeClick-Fetch', 'Medium', 'Config', e.message || JSON.stringify(e));
                  }
                }
                this.context.propertyPane.refresh();
              };
              const hasCascade = this.properties.cascadeConfig && this.properties.cascadeConfig[field.key];
              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(isConfiguringCascade || !!hasCascade),
                  onClick: onCascadeClick,
                  title: "Cascade Configuration"
                }, React.createElement(Icons.Cascade))
              );
            }

            // 6. Autocomplete Button
            if ((type === 'add' || type === 'edit') && field.type === 'Text') {
              const onConfigClick = async () => {
                try {
                  if (isConfiguringAutocomplete) {
                    this.activeAutocompleteField = null;
                    this.sourceFields = [];
                  } else {
                    this.activeLookupConfigField = null;
                    this.activeCascadeField = null;
                    this.properties.validationEditorFieldKey = undefined;
                    this.properties.permissionEditorFieldKey = undefined;
                    this.activeAutocompleteField = field.key;
                    if (!this.lists || this.lists.length === 0) { await this.loadLists(); }
                    this.sourceLists = [...this.lists];
                    const currentCfg = (this.properties.autocompleteConfig || {})[field.key];
                    if (currentCfg && currentCfg.sourceList) {
                      await this.loadSourceFields(currentCfg.sourceList);
                    }
                  }
                  this.context.propertyPane.refresh();
                } catch (error: any) {
                  void LoggerService.log('PowerFormWebPart-onConfigClick', 'Medium', 'Config', error.message || JSON.stringify(error));
                }
              };
              const hasAC = this.properties.autocompleteConfig && this.properties.autocompleteConfig[field.key] && this.properties.autocompleteConfig[field.key].sourceList;
              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(isConfiguringAutocomplete || !!hasAC),
                  onClick: onConfigClick,
                  title: "Autocomplete Configuration"
                }, React.createElement(Icons.Autocomplete))
              );
            }

            // 7. Filter Toggle (List Mode)
            if (type === 'list') {
              const isFilterOn = this.properties.listFilterMap && this.properties.listFilterMap[field.key] === true;
              const onToggleFilter = () => {
                const fMap = this.properties.listFilterMap || {};
                if (fMap[field.key]) delete fMap[field.key];
                else fMap[field.key] = true;
                this.properties.listFilterMap = fMap;
                this.context.propertyPane.refresh();
              };
              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(isFilterOn, '#107c10'),
                  onClick: onToggleFilter,
                  title: isFilterOn ? "Filter: ON" : "Filter: OFF"
                }, isFilterOn ? React.createElement(Icons.Filter) : React.createElement(Icons.FilterOff))
              );
            }

            // 8. Group By Choice Toggle
            if (type === 'list' && field.type === 'Choice') {
              const currentGroupField = this.properties.listGroupByField;
              const isGrouped = currentGroupField === field.key;
              const isGroupLocked = !!currentGroupField && currentGroupField !== field.key;

              const onToggleGroup = () => {
                this.properties.listGroupByField = isGrouped ? '' : field.key;
                this.context.propertyPane.refresh();
              };

              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(isGrouped, '#0078d4'),
                  onClick: onToggleGroup,
                  disabled: isGroupLocked,
                  title: isGrouped ? "Ungroup this column" : (isGroupLocked ? "Disable the existing Group first" : "Group by this column")
                }, React.createElement(Icons.Group))
              );
            }

            // 9. Conditional Formatting
            if (type === 'list' && (field.type === 'Choice' || field.type === 'MultiChoice' || field.type === 'DateTime')) {
              const isFormattingActive = this.properties['activeFormattingField'] === field.key;
              const hasConfig = this.properties.formattingConfig && this.properties.formattingConfig[field.key];
              const onFormattingClick = () => {
                this.properties['activeFormattingField'] = isFormattingActive ? undefined : field.key;
                this.context.propertyPane.refresh();
              };
              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(isFormattingActive || !!hasConfig, '#008272'),
                  onClick: onFormattingClick,
                  title: "Conditional Formatting"
                }, React.createElement("svg", { style: iconStyle, viewBox: "0 0 24 24" }, React.createElement("path", { d: "M12 2.69l5.66 5.66a8 8 0 1 1-11.31 0z" })))
              );
            }

            // 10. Repeater Grid
            if (field.type === 'Note' && (type === 'add' || type === 'edit')) {
              const isRepeaterConfigured = this.properties.repeaterConfig && this.properties.repeaterConfig[field.key];
              const onConfigRepeater = () => {
                this.activeRepeaterFieldKey = field.key;
                this.context.propertyPane.refresh();
              };
              children.push(
                React.createElement('button', {
                  style: iconBtnStyle(!!isRepeaterConfigured, '#d13438'),
                  onClick: onConfigRepeater,
                  title: "Configure Repeater Grid"
                }, React.createElement(Icons.Grid))
              );
            }

            const container = React.createElement('div', { style: { display: 'flex', alignItems: 'center' } }, ...children);
            ReactDom.render(container, elem);
          } catch (error: any) {
            void LoggerService.log('PowerFormWebPart-renderFieldRow', 'Medium', 'Config', error.message || JSON.stringify(error));
          }
        }
      });

      // Formatting Editor Section
      if (type === 'list' && this.properties.activeFormattingField === field.key) {
        formattingEditor = PropertyPaneCustomField({
          key: `fmt_editor_${field.key}`,
          onRender: (elem) => {
            const editor = React.createElement(FormattingEditor, {
              field: field as any,
              config: (this.properties.formattingConfig && this.properties.formattingConfig[field.key]) || {
                type: field.type === 'DateTime' ? 'date' : 'choice',
                choiceConfig: {},
                dateRules: []
              },
              onSave: (newCfg) => {
                if (!this.properties.formattingConfig) this.properties.formattingConfig = {};
                this.properties.formattingConfig[field.key] = newCfg;
                this.properties.activeFormattingField = undefined;
                this.context.propertyPane.refresh();
              },
              onCancel: () => {
                this.properties.activeFormattingField = undefined;
                this.context.propertyPane.refresh();
              }
            });
            ReactDom.render(editor, elem);
          }
        });
      }

      // --- RENDER DYNAMIC CONFIG PANELS ---
      const autocompleteControls: IPropertyPaneField<any>[] = [];
      if ((type === 'add' || type === 'edit') && isConfiguringAutocomplete) {
        if (!this.properties.autocompleteConfig) this.properties.autocompleteConfig = {};
        if (!this.properties.autocompleteConfig[field.key]) this.properties.autocompleteConfig[field.key] = { sourceList: '', sourceField: '' };

        autocompleteControls.push(
          PropertyPaneDropdown(`autocompleteConfig[${field.key}].sourceList`, { label: '1. Source List', options: this.sourceLists, selectedKey: this.properties.autocompleteConfig[field.key].sourceList }),
          PropertyPaneDropdown(`autocompleteConfig[${field.key}].sourceField`, { label: '2. Main Display Column', options: this.sourceFields, selectedKey: this.properties.autocompleteConfig[field.key].sourceField, disabled: this.sourceFields.length === 0 }),
          PropertyPaneCustomField({
            key: `ac_multi_${field.key}`,
            onRender: (dom) => {
              const currentAdditional = this.properties.autocompleteConfig[field.key].additionalFields || [];
              const element = React.createElement(CheckboxListEditor, {
                label: "3. Additional Display Columns",
                options: this.sourceFields.map(f => ({ key: f.key, text: f.text })),
                selectedKeys: currentAdditional,
                maxSelection: 5,
                onChanged: (newKeys: string[]) => {
                  this.properties.autocompleteConfig[field.key].additionalFields = newKeys;
                  this.render();
                }
              });
              ReactDom.render(element, dom);
            }
          }),
          PropertyPaneTextField(`autocompleteConfig[${field.key}].sourceQuery`, { label: '4. Optional Filter Query (OData)', placeholder: "e.g. Status eq 'Active'" }),
          PropertyPaneCustomField({
            key: `ac_map_${field.key}`,
            onRender: (dom) => {
              const currentMap = this.properties.autocompleteConfig[field.key].columnMapping || [];
              const element = React.createElement(ColumnMappingEditor, {
                label: "5. Map Columns (Auto-Populate)",
                sourceOptions: this.sourceFields,
                targetOptions: this.fields.map(f => ({ key: f.key, text: f.text })),
                mappings: currentMap,
                onChanged: (newMap: IColumnMapping[]) => {
                  this.properties.autocompleteConfig[field.key].columnMapping = newMap;
                  this.render();
                }
              });
              ReactDom.render(element, dom);
            }
          }),
          PropertyPaneButton('apply_ac', {
            text: 'Refresh Source Columns', buttonType: PropertyPaneButtonType.Normal, icon: 'Refresh',
            onClick: async () => {
              try {
                const list = this.properties.autocompleteConfig[field.key].sourceList;
                if (list) await this.loadSourceFields(list);
                this.context.propertyPane.refresh();
              } catch (error: any) {
                void LoggerService.log('PowerFormWebPart-RefreshAC', 'Low', 'Config', error.message);
              }
            }
          }),
          PropertyPaneButton('clear_ac', {
            text: 'Clear Configuration', buttonType: PropertyPaneButtonType.Normal, icon: 'Delete',
            onClick: () => {
              const newConfig = JSON.parse(JSON.stringify(this.properties.autocompleteConfig));
              delete newConfig[field.key];
              this.properties.autocompleteConfig = newConfig;
              this.activeAutocompleteField = null;
              this.context.propertyPane.refresh();
            }
          })
        );
      }

      const lookupControls: IPropertyPaneField<any>[] = [];
      if ((type === 'add' || type === 'edit') && this.activeLookupConfigField === field.key) {
        if (!this.properties.lookupDisplayConfig) this.properties.lookupDisplayConfig = {};
        if (!this.properties.lookupDisplayConfig[field.key]) this.properties.lookupDisplayConfig[field.key] = { additionalFields: [] };

        lookupControls.push(
          PropertyPaneCustomField({
            key: `lu_multi_${field.key}`,
            onRender: (dom) => {
              const config = this.properties.lookupDisplayConfig;
              const currentAdditional = (config && config[field.key] && config[field.key].additionalFields) ? config[field.key].additionalFields : [];
              const element = React.createElement(CheckboxListEditor, {
                label: "Additional Display Columns (Max 5)",
                options: this.sourceFields.map(f => ({ key: f.key, text: f.text })),
                selectedKeys: currentAdditional,
                maxSelection: 5,
                onChanged: (newKeys: string[]) => {
                  if (!this.properties.lookupDisplayConfig) this.properties.lookupDisplayConfig = {};
                  if (!this.properties.lookupDisplayConfig[field.key]) this.properties.lookupDisplayConfig[field.key] = { additionalFields: [] };
                  this.properties.lookupDisplayConfig[field.key].additionalFields = newKeys;
                  this.render();
                }
              });
              ReactDom.render(element, dom);
            }
          }),
          PropertyPaneTextField(`lookupDisplayConfig[${field.key}].filterQuery`, { label: 'Optional Filter Query (OData)', placeholder: "e.g. Status eq 'Active'" })
        );

        if (field.type === 'Lookup') {
          lookupControls.push(PropertyPaneCustomField({
            key: `lu_map_${field.key}`,
            onRender: (dom) => {
              const config = this.properties.lookupDisplayConfig;
              const currentMap = (config && config[field.key] && config[field.key].columnMapping) ? config[field.key].columnMapping : [];
              const element = React.createElement(ColumnMappingEditor, {
                label: "Map Columns (Auto-Populate)",
                sourceOptions: this.sourceFields,
                targetOptions: this.fields.map(f => ({ key: f.key, text: f.text })),
                mappings: currentMap,
                onChanged: (newMap: IColumnMapping[]) => {
                  if (!this.properties.lookupDisplayConfig) this.properties.lookupDisplayConfig = {};
                  if (!this.properties.lookupDisplayConfig[field.key]) this.properties.lookupDisplayConfig[field.key] = { additionalFields: [], columnMapping: [] };
                  this.properties.lookupDisplayConfig[field.key].columnMapping = newMap;
                  this.render();
                }
              });
              ReactDom.render(element, dom);
            }
          }));
        }

        lookupControls.push(
          PropertyPaneButton(`lu_refresh_${field.key}`, {
            text: 'Refresh Columns', buttonType: PropertyPaneButtonType.Normal, icon: 'Refresh',
            onClick: async () => {
              const fDefUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/fields/getByInternalNameOrTitle('${field.key}')`;
              const fRes = await this.context.spHttpClient.get(fDefUrl, SPHttpClient.configurations.v1);
              const fData = await fRes.json();
              if (fData.LookupList) { await this.loadSourceFields(fData.LookupList); }
              this.context.propertyPane.refresh();
            }
          }),
          PropertyPaneButton(`lu_clear_${field.key}`, {
            text: 'Clear Configuration', buttonType: PropertyPaneButtonType.Normal, icon: 'Delete',
            onClick: () => {
              const newConfig = JSON.parse(JSON.stringify(this.properties.lookupDisplayConfig));
              delete newConfig[field.key];
              this.properties.lookupDisplayConfig = newConfig;
              this.activeLookupConfigField = null;
              this.context.propertyPane.refresh();
            }
          })
        );
      }

      const cascadeControls: IPropertyPaneField<any>[] = [];
      if (isConfiguringCascade) {
        if (!this.properties.cascadeConfig) this.properties.cascadeConfig = {};
        if (!this.properties.cascadeConfig[field.key]) { this.properties.cascadeConfig[field.key] = { parentField: '', foreignKey: '' }; }
        const parentOptions = this.fields.filter(f => f.key !== field.key).map(f => ({ key: f.key, text: f.text }));

        cascadeControls.push(
          PropertyPaneDropdown(`cascadeConfig[${field.key}].parentField`, { label: 'Parent Field (Source)', options: parentOptions, selectedKey: this.properties.cascadeConfig[field.key].parentField }),
          PropertyPaneTextField(`cascadeConfig[${field.key}].foreignKey`, { label: 'Foreign Key Column (Internal Name)', description: 'Internal name of column in child list pointing to Parent.', placeholder: 'e.g. CountryRef' }),
          PropertyPaneCustomField({
            key: `cas_multi_${field.key}`,
            onRender: (dom) => {
              const currentAdditional = this.properties.cascadeConfig?.[field.key]?.additionalFields || [];
              const element = React.createElement(CheckboxListEditor, {
                label: "Additional Display Columns (Max 5)",
                options: this.sourceFields.map(f => ({ key: f.key, text: f.text })),
                selectedKeys: currentAdditional,
                maxSelection: 5,
                onChanged: (newKeys: string[]) => {
                  if (this.properties.cascadeConfig && this.properties.cascadeConfig[field.key]) {
                    this.properties.cascadeConfig[field.key].additionalFields = newKeys;

                  }
                  this.render();
                }
              });
              ReactDom.render(element, dom);
            }
          }),
          PropertyPaneTextField(`cascadeConfig[${field.key}].filterQuery`, { label: 'Optional Filter Query (OData)', placeholder: "e.g. IsActive eq 1" }),
          PropertyPaneCustomField({
            key: `cas_map_${field.key}`,
            onRender: (dom) => {
              const currentMap = this.properties.cascadeConfig?.[field.key]?.columnMapping || [];
              const element = React.createElement(ColumnMappingEditor, {
                label: "Map Columns (Auto-Populate)",
                sourceOptions: this.sourceFields,
                targetOptions: this.fields.map(f => ({ key: f.key, text: f.text })),
                mappings: currentMap,
                onChanged: (newMap: IColumnMapping[]) => {
                  if (this.properties.cascadeConfig && this.properties.cascadeConfig[field.key]) {

                    this.properties.cascadeConfig[field.key].columnMapping = newMap;
                  }
                  this.render();
                }
              });
              ReactDom.render(element, dom);
            }
          }),
          PropertyPaneButton(`cas_refresh_${field.key}`, {
            text: 'Refresh Source Columns', buttonType: PropertyPaneButtonType.Normal, icon: 'Refresh',
            onClick: async () => {
              const fDefUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/fields/getByInternalNameOrTitle('${field.key}')`;
              const fRes = await this.context.spHttpClient.get(fDefUrl, SPHttpClient.configurations.v1);
              const fData = await fRes.json();
              if (fData.LookupList) { await this.loadSourceFields(fData.LookupList); }
              this.context.propertyPane.refresh();
            }
          }),
          PropertyPaneButton(`clear_cascade_${field.key}`, {
            text: 'Remove Cascade Configuration', buttonType: PropertyPaneButtonType.Normal, icon: 'Delete',
            onClick: () => {
              const newConfig = JSON.parse(JSON.stringify(this.properties.cascadeConfig));
              delete newConfig[field.key];
              this.properties.cascadeConfig = newConfig;
              this.activeCascadeField = null;
              this.context.propertyPane.refresh();
            }
          })
        );
      }

      let editor = null;
      if (isEditingValidation) {
        const fieldRules = (this.properties.validationConfig && this.properties.validationConfig[field.key]) || [];
        editor = PropertyPaneCustomField({
          key: 'edit_val_' + field.key,
          onRender: (elem: HTMLElement) => {
            const onSave = (rules: ICustomValidationRule[]) => {
              const cfg = cloneConfig(this.properties.validationConfig);
              cfg[field.key] = rules;
              this.properties.validationConfig = cfg;
              this.properties.validationEditorFieldKey = undefined;
              this.context.propertyPane.refresh();
            };
            const onCancel = () => {
              this.properties.validationEditorFieldKey = undefined;
              this.context.propertyPane.refresh();
            };
            const cmp = React.createElement(ValidationEditor as any, { fieldKey: field.key, rules: fieldRules, onSave, onCancel } as any);
            ReactDom.render(cmp, elem);
          }
        });
      }

      let permEditor = null;
      if (isEditingPerms) {
        const currentGroups = (this.properties.fieldPermissionConfig && this.properties.fieldPermissionConfig[field.key]) || [];
        permEditor = PropertyPaneCustomField({
          key: 'perm_edit_' + field.key,
          onRender: (elem: HTMLElement) => {
            const onSave = (groups: string[]) => {
              if (!this.properties.fieldPermissionConfig) this.properties.fieldPermissionConfig = {};
              this.properties.fieldPermissionConfig[field.key] = groups;
              this.properties.permissionEditorFieldKey = undefined;
              this.context.propertyPane.refresh();
            };
            const onCancel = () => {
              this.properties.permissionEditorFieldKey = undefined;
              this.context.propertyPane.refresh();
            };
            const cmp = React.createElement(FieldPermissionEditor, { fieldKey: field.key, fieldTitle: field.text, context: this.context, selectedGroups: currentGroups, onSave, onCancel });
            ReactDom.render(cmp, elem);
          }
        });
      }

      const result = [row, ...autocompleteControls, ...lookupControls, ...cascadeControls];
      if (editor) result.push(editor);
      if (permEditor) result.push(permEditor);
      if (formattingEditor) result.push(formattingEditor);

      if (this.activeRepeaterFieldKey === field.key) {
        const currentCfg = (this.properties.repeaterConfig && this.properties.repeaterConfig[field.key]) || [];
        const repEditor = PropertyPaneCustomField({
          key: `rep_edit_${field.key}`,
          onRender: (elem: HTMLElement) => {
            const component = React.createElement(RepeaterConfigEditor, {
              fieldKey: field.key, currentConfig: currentCfg,
              onSave: (newCfg) => {
                if (!this.properties.repeaterConfig) this.properties.repeaterConfig = {};
                this.properties.repeaterConfig[field.key] = newCfg;
                this.activeRepeaterFieldKey = null;
                this.context.propertyPane.refresh();
              },
              onCancel: () => { this.activeRepeaterFieldKey = null; this.context.propertyPane.refresh(); },
              onClear: () => {
                if (confirm('Remove repeater configuration? Field will revert to standard input.')) {
                  if (this.properties.repeaterConfig && this.properties.repeaterConfig[field.key]) { delete this.properties.repeaterConfig[field.key]; }
                  this.activeRepeaterFieldKey = null;
                  this.context.propertyPane.refresh();
                }
              }
            });
            ReactDom.render(component, elem);
          }
        });
        result.push(repEditor);
      }

      return result;
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-getFieldControl', 'High', 'Config', error.message || JSON.stringify(error));
      return [PropertyPaneLabel('err', { text: 'Err' }) as any];
    }
  }

  // --- CONFIGURATION IMPORT/EXPORT LOGIC ---
  /**
   * Triggers a browser download of the current web part properties as a JSON file.
   */
  private exportConfiguration = (): void => {
    try {
      // 1. Get current properties
      const configToExport = JSON.parse(JSON.stringify(this.properties));
      // Optional: Clean up internal UI state flags if you don't want them persisted/exported
      delete configToExport.isConfiguringEditActions;
      delete configToExport.isConfiguringViewActions;
      delete configToExport.isConfiguringDefaultViewPerms;
      delete configToExport.isConfiguringViews;
      delete configToExport.validationEditorFieldKey;
      delete configToExport.permissionEditorFieldKey;
      delete configToExport.autocompleteEditorFieldKey;
      // 2. Create Blob
      const dataStr = JSON.stringify(configToExport, null, 2);
      const dataUri = 'data:application/json;charset=utf-8,' + encodeURIComponent(dataStr);
      // 3. Trigger Download
      const exportFileDefaultName = `PowerForm_Config_${this.properties.selectedList || 'Export'}_${new Date().toISOString().slice(0, 10)}.json`;
      const linkElement = document.createElement('a');
      linkElement.setAttribute('href', dataUri);
      linkElement.setAttribute('download', exportFileDefaultName);
      linkElement.click();
    } catch (e: any) {
      void LoggerService.log('PowerFormWebPart-exportConfiguration', 'High', 'Config', e.message || JSON.stringify(e));
      void Swal.fire('Error!', 'Failed to export configuration.');
    }
  }
  /**
   * Triggers the hidden file input to open the browser file picker.
   */
  private triggerImport = (): void => {
    try {
      const fileInput = document.getElementById('config-import-input') as HTMLInputElement;
      if (fileInput) fileInput.click();
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-triggerImport', 'Low', 'Config', error.message || JSON.stringify(error));
    }
  }
  /**
   * Reads the selected JSON file and applies it to this.properties.
   */
  private handleImportFile = (event: any): void => {
    try {
      const file = event.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (e: any) => {
        try {
          const importedConfig = JSON.parse(e.target.result);
          // Basic Validation: Check if it looks like our config (has selectedList or specific arrays)
          if (importedConfig.selectedList === undefined && importedConfig.addVisibleFields === undefined) {
            void Swal.fire('Error!', 'Invalid configuration file. Missing standard properties.');
            return;
          }
          if (confirm('This will overwrite your current configuration. Are you sure?')) {
            // Apply Configuration
            // We assign keys individually or use Object.assign to merge into the existing properties object
            // This ensures the SPFx web part infrastructure detects the change.
            (Object as any).assign(this.properties, importedConfig);
            // Refresh UI
            this.context.propertyPane.refresh();
            this.render();
            // If the list changed, we might want to reload fields
            if (importedConfig.selectedList) {
              void this.loadFieldsForSelectedList(importedConfig.selectedList).then(() => {
                this.context.propertyPane.refresh();
              });
            }
            void Swal.fire({ icon: 'success', title: 'success', text: "Configuration imported successfully!" });
          }
        } catch (err: any) {
          void LoggerService.log('PowerFormWebPart-handleImportFile-Parser', 'High', 'Config', err.message || JSON.stringify(err));
          void Swal.fire('Error!', 'Error parsing JSON file.');
        }
        // Reset input so same file can be selected again if needed
        event.target.value = '';
      };
      reader.readAsText(file);
    } catch (error: any) {
      void LoggerService.log('PowerFormWebPart-handleImportFile', 'High', 'Config', error.message || JSON.stringify(error));
    }
  }
  private async provisionLogEnvironment(): Promise<void> {
    const { logListTitle, logRoleName } = this.properties;
    if (!logListTitle || !logRoleName) {
      void Swal.fire({ icon: 'info', title: 'Missing Info', text: "Please enter both a Log List Name and a Role Name." });
      return;
    }

    try {
      // --- SHOW LOADING POPUP ---
      void Swal.fire({
        title: 'Provisioning Logs...',
        html: `Creating list <b>${logListTitle}</b> and configuring permissions. Please wait...`,
        allowOutsideClick: false,
        didOpen: () => { Swal.showLoading(); }
      });
      const web = this._sp.web;

      // Check if list exists
      try {
        await web.lists.getByTitle(logListTitle)();
        void Swal.fire({ icon: 'info', title: 'Exists', text: `List "${logListTitle}" already exists.` });
      } catch (e: any) {
        // 1. Create List
        const listResult = await web.lists.add(logListTitle, "System Logs", 100);
        const list = web.lists.getByTitle(logListTitle);

        // 2. Add Fields
        await list.fields.addText("Page");
        await list.fields.addText("ItemId");
        await list.fields.addText("Module");
        await list.fields.addText("Severity");
        await list.fields.addMultilineText("Error");
        await list.fields.addText("ErrorId");

        // 3. Add Fields to Default View
        try {
          const view = list.views.getByTitle("All Items");
          const fieldsToShow = ["Page", "ItemId", "Module", "Severity", "Error", "ErrorId"];
          for (const field of fieldsToShow) {
            try { await view.fields.add(field); } catch (viewErr) { }
          }
        } catch (viewFetchErr) {
          console.error("Could not fetch 'All Items' view to update columns.", viewFetchErr);
        }

        // 4. Permissions
        await list.breakRoleInheritance(false, true);
        try {
          const roleDef = await web.roleDefinitions.getByName(logRoleName)();
          const roleDefId = roleDef.Id;
          const everyoneIdentifier = "c:0(.s|true";
          const everyoneUser = await web.ensureUser(everyoneIdentifier);
          await list.roleAssignments.add(everyoneUser.Id, roleDefId);
        } catch (permErr) {
          console.warn("Could not set permission automatically", permErr);
        }

        void Swal.fire({ icon: 'success', title: 'success', text: `Log List "${logListTitle}" created successfully.` });
      }
    } catch (e: any) {
      console.error("Critical Error in provisionLogEnvironment:", e);
      void Swal.fire({ icon: 'error', title: 'Error', text: e.message });
    }
  }

  // Helper to generate a single rule field set
  private getSingleRuleConfig(prefix: string, index: number): any[] {
    const keyPrefix = `${prefix}${index}`; // e.g., ruleAdd1
    const isEnabled = (this.properties as any)[`${keyPrefix}Enable`];

    return [
      PropertyPaneLabel(`${keyPrefix}_lbl`, { text: `--- Rule #${index} ---` }),
      PropertyPaneToggle(`${keyPrefix}Enable`, { label: "Enable" }),
      PropertyPaneTextField(`${keyPrefix}Cond`, {
        label: "Condition (OData)",
        placeholder: "Status eq 'Approved'",
        disabled: !isEnabled
      }),
      PropertyPaneTextField(`${keyPrefix}Msg`, {
        label: "Message",
        multiline: true,
        disabled: !isEnabled
      }),
      this.getGroupPickerControl('Target Groups (Empty = Everyone)', `${keyPrefix}Groups`)
    ];
  }

  // Helper to get all rules for a category (hidden behind a toggle)
  private getCategoryRulesGroup(categoryName: string, prefix: string, showPropName: string): any {
    const showRules = (this.properties as any)[showPropName];

    const fields: any[] = [
      PropertyPaneToggle(showPropName, {
        label: `Configure Advanced Rules (${categoryName})`,
        onText: "Show Rules",
        offText: "Hide Rules"
      })
    ];

    if (showRules) {
      // Add 5 rules
      for (let i = 1; i <= 5; i++) {
        fields.push(...this.getSingleRuleConfig(prefix, i));
      }
    }

    return fields;
  }

  private getRulePageGroup(index: number): any {
    const i = index;
    return {
      groupName: `Rule #${i}`,
      groupFields: [
        PropertyPaneToggle(`rule${i}Enable`, { label: "Enable Rule" }),
        PropertyPaneTextField(`rule${i}Cond`, {
          label: "Condition (OData)",
          placeholder: "Status eq 'Approved'",
          disabled: !(this.properties as any)[`rule${i}Enable`]
        }),
        PropertyPaneTextField(`rule${i}Msg`, {
          label: "Message",
          multiline: true,
          disabled: !(this.properties as any)[`rule${i}Enable`]
        }),
        this.getGroupPickerControl('Target Groups', `rule${i}Groups`)
      ]
    };
  }
  private getGroupPickerControl(label: string, propertyName: string): IPropertyPaneField<any> {
    return PropertyPaneCustomField({
      key: `grp_picker_${propertyName}`,
      onRender: (elem) => {
        // Safe access to the property
        const currentKeys = (this.properties as any)[propertyName] || [];

        // Ensure keys are strings for comparison in the editor
        const selectedKeys = currentKeys.map((k: any) => String(k.id || k));

        // Render the CheckboxListEditor (Groups Only)
        const element = React.createElement(CheckboxListEditor, {
          label: label,
          options: this.siteGroups, // Uses the groups loaded in onInit
          selectedKeys: selectedKeys,
          onChanged: (newKeys: string[]) => {
            // Save as array of Numbers (Group IDs)
            (this.properties as any)[propertyName] = newKeys.map(k => parseInt(k, 10));
            this.context.propertyPane.refresh();
          }
        });
        ReactDom.render(element, elem);
      }
    });
  }

  private getCustomRuleEditorControl(label: string, propertyName: string, stateFlagName: string): IPropertyPaneField<any> {
    const isEditing = (this.properties as any)[stateFlagName];
    const currentRules = (this.properties as any)[propertyName] || [];

    if (isEditing) {
      return PropertyPaneCustomField({
        key: `editor_${propertyName}`,
        onRender: (elem) => {
          const editor = React.createElement(NotificationRuleEditor, {
            label: label,
            rules: currentRules,
            siteGroups: this.siteGroups,
            onSave: (newRules: any[]) => {
              (this.properties as any)[propertyName] = newRules;
              (this.properties as any)[stateFlagName] = false;
              this.context.propertyPane.refresh();
            },
            onCancel: () => {
              (this.properties as any)[stateFlagName] = false;
              this.context.propertyPane.refresh();
            }
          });
          ReactDom.render(editor, elem);
        }
      });
    }

    return PropertyPaneCustomField({
      key: `btn_${propertyName}`,
      onRender: (elem) => {
        const btn = React.createElement('button', {
          onClick: () => {
            (this.properties as any)[stateFlagName] = true;
            this.context.propertyPane.refresh();
          },
          style: { padding: '8px 16px', cursor: 'pointer', backgroundColor: '#fff', border: '1px solid #0078d4', color: '#0078d4', borderRadius: '4px', width: '100%', marginTop: '10px' }
        }, `Configure ${label} (${currentRules.length})`);
        ReactDom.render(btn, elem);
      }
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    try {
      // 1. DETERMINE STATE
      const isFirstTime = !this.properties.isLogSetupConfirmed;
      let maxOrder = this.fields.filter(f => f.key !== 'ContentType').length;
      // ==============================================================================
      // PAGE 1: MANDATORY LOGGING SETUP (From getPropertyPaneConfiguration)
      // ==============================================================================
      const setupPage: IPropertyPanePage = {
        header: { description: isFirstTime ? 'Environment Setup' : 'Logging Configuration' },
        groups: [{
          groupName: 'Step 1: Logging Infrastructure',
          groupFields: [
            PropertyPaneChoiceGroup('installType', {
              label: 'Is this the first time configuring logs in this site?',
              options: [
                { key: 'new', text: 'Yes, create new environment', iconProps: { officeFabricIconFontName: 'Add' } },
                { key: 'existing', text: 'No, use existing log list', iconProps: { officeFabricIconFontName: 'Link' } }
              ]
            }),
            // Fields for NEW installation
            ...(this.properties.installType === 'new' ? [
              PropertyPaneTextField('logListTitle', { label: 'New Log List Name:', placeholder: 'e.g. App_Logs' }),
              PropertyPaneTextField('logRoleName', { label: 'Add-Only Role Name:', placeholder: 'e.g. Logs_AddOnly_Role' }),
              PropertyPaneButton('btnProvision', {
                text: 'Automate: Create List & Set Security',
                buttonType: PropertyPaneButtonType.Primary,
                onClick: this.provisionLogEnvironment.bind(this)
              })
            ] : []),
            // Fields for EXISTING installation
            ...(this.properties.installType === 'existing' ? [
              PropertyPaneDropdown('logListTitle', { label: 'Select Existing Log List:', options: this.lists })
            ] : []),
            PropertyPaneToggle('isLogSetupConfirmed', {
              label: 'Confirm Setup to Unlock Form Settings',
              checked: this.properties.isLogSetupConfirmed,
              disabled: !this.properties.logListTitle
            })
          ]
        }]
      };
      // ==============================================================================
      // HELPER: FIELD GROUPS GENERATOR (From getPropertyPaneConfiguration)
      // ==============================================================================
      const getGroup = (type: 'add' | 'edit' | 'view' | 'list', title: string) => {
        let fields: IPropertyPaneField<IPropertyPaneCustomFieldProps>[] = [];
        // 1. ADD SELECT ALL BUTTON
        fields.push(this.getSelectAllControl(type));
        // 2. ADD FIELDS
        for (let i = 0; i < this.fields.length; i++) {
          const f = this.fields[i];
          if (f.key === 'ContentType') continue;
          const isSystemAuditField = ['Created', 'Author', 'Modified', 'Editor'].indexOf(f.key) > -1;
          if (type !== 'list' && isSystemAuditField) {
            continue;
          }
          let ctrls = this.getFieldControl(f, maxOrder, type);
          fields.push(...ctrls);
        }
        // 3. ADD MODE SPECIFIC BUTTONS
        if (type === 'edit') {
          fields.push(this.getActionConfigControl('edit'));
        } else if (type === 'view') {
          fields.push(this.getActionConfigControl('view'));
        }
        if (type === 'list') {
          // VIEW EDITOR Logic
          fields.push(
            PropertyPaneCustomField({
              key: 'viewEditorConfig',
              onRender: (elem) => {
                try {
                  const isConfiguring = this.properties['isConfiguringViews'];
                  if (isConfiguring) {
                    const editor = React.createElement(ViewEditor, {
                      views: this.properties.views || [],
                      fields: this.fields,
                      context: this.context,
                      onSave: (newViews) => {
                        this.properties.views = newViews;
                        this.properties['isConfiguringViews'] = false;
                        this.context.propertyPane.refresh();
                      },
                      onCancel: () => {
                        this.properties['isConfiguringViews'] = false;
                        this.context.propertyPane.refresh();
                      }
                    });
                    ReactDom.render(editor, elem);
                  } else {
                    const btn = React.createElement('button', {
                      onClick: () => {
                        this.properties['isConfiguringViews'] = true;
                        this.context.propertyPane.refresh();
                      },
                      style: {
                        padding: '8px 16px', cursor: 'pointer', backgroundColor: '#fff',
                        border: '1px solid #0078d4', color: '#0078d4', borderRadius: '4px', width: '100%', marginBottom: '10px'
                      }
                    }, `Configure Views (${(this.properties.views || []).length})`);
                    ReactDom.render(btn, elem);
                  }
                } catch (error: any) {
                  void LoggerService.log('PowerFormWebPart-renderViewEditor', 'Medium', 'Config', error.message);
                }
              }
            })
          );
          // DEFAULT VIEW PERMISSIONS Logic
          if (this.properties.isConfiguringDefaultViewPerms) {
            fields.push(PropertyPaneCustomField({
              key: 'def_view_perm',
              onRender: (elem) => {
                try {
                  const cmp = React.createElement(FieldPermissionEditor, {
                    fieldKey: 'DefaultView',
                    fieldTitle: 'Default View (All Items)',
                    context: this.context,
                    selectedGroups: this.properties.defaultViewAllowedGroups || [],
                    onSave: (groups: string[]) => {
                      this.properties.defaultViewAllowedGroups = groups;
                      this.properties.isConfiguringDefaultViewPerms = false;
                      this.context.propertyPane.refresh();
                    },
                    onCancel: () => {
                      this.properties.isConfiguringDefaultViewPerms = false;
                      this.context.propertyPane.refresh();
                    }
                  });
                  ReactDom.render(cmp, elem);
                } catch (error: any) {
                  void LoggerService.log('PowerFormWebPart-renderDefViewPerms', 'Medium', 'Config', error.message);
                }
              }
            }));
          } else {
            const count = (this.properties.defaultViewAllowedGroups || []).length;
            fields.push(PropertyPaneButton('btn_def_view_perm', {
              text: `Set Default View Permissions (${count > 0 ? count + ' Groups' : 'Everyone'})`,
              icon: 'Permissions',
              onClick: () => {
                this.properties.isConfiguringDefaultViewPerms = true;
                this.context.propertyPane.refresh();
              }
            }) as any);
          }
        }
        // 4. CUSTOM SCRIPT
        const scriptProp = type === 'add' ? 'addCustomScript' :
          type === 'edit' ? 'editCustomScript' :
            type === 'view' ? 'viewCustomScript' :
              'listCustomScript';
        fields.push(
          PropertyPaneTextField(scriptProp, {
            label: 'Custom JavaScript (Run on Load) - Additional Customizations',
            multiline: true,
            rows: 6,
            placeholder: 'Enter URL (https://...js) or plain JS code here.'
          }) as any
        );
        const styleProp = type === 'add' ? 'addCustomStyle' :
          type === 'edit' ? 'editCustomStyle' :
            type === 'view' ? 'viewCustomStyle' :
              'listCustomStyle';

        fields.push(
          PropertyPaneTextField(styleProp, {
            label: 'Custom CSS (Styles) - Additional Styling',
            multiline: true,
            rows: 6,
            placeholder: 'Enter URL (https://...css) or plain CSS code (e.g. .ms-Label { color: red !important; }) here.'
          }) as any
        );
        // 5. FINAL STEP: Return Header and Groups including the shared Action Buttons  
        return {
          header: { description: title },
          groups: [{ groupName: 'Fields', groupFields: fields }, (this as any).getConfigurationActionGroup()]
        };
      };
      // ==============================================================================
      // DYNAMIC PAGE ASSEMBLY
      // ==============================================================================
      // Start with Setup Page
      const pages: IPropertyPanePage[] = [setupPage];
      // Only add the full feature set if Setup is Confirmed  
      if (this.properties.isLogSetupConfirmed) {
        pages.push(
          // --- 1. GENERAL SETTINGS (From Hold) ---
          {
            header: { description: 'General Settings' },
            displayGroupsAsAccordion: true,
            groups: [
              {
                groupName: 'Main Configuration',
                groupFields: [
                  PropertyPaneCustomField({
                    key: 'importConfig',
                    onRender: (elem) => {
                      const container = React.createElement('div', {},
                        React.createElement('input', {
                          id: 'config-import-input', type: 'file', accept: '.json', style: { display: 'none' },
                          onChange: this.handleImportFile.bind(this)
                        }),
                        React.createElement('button', {
                          onClick: this.triggerImport,
                          style: {
                            width: '100%', padding: '8px', marginBottom: '15px', backgroundColor: '#f3f2f1',
                            border: '1px dashed #8a8886', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', fontWeight: 600
                          },
                          title: 'Upload a .json configuration file'
                        }, React.createElement('span', { style: { fontSize: '16px' } }, '📂'), 'Import Configuration')
                      );
                      ReactDom.render(container, elem);
                    }
                  }),
                  PropertyPaneToggle('showHiddenLists', {
                    label: 'Show Hidden System Lists',
                    onText: 'Show All (Template 100)',
                    offText: 'Standard Only',
                    checked: this.properties.showHiddenLists
                  }),
                  PropertyPaneDropdown('selectedList', { label: 'List', options: this.lists }),
                  PropertyPaneTextField('listPageTitle', { label: 'Listing Screen Title' }),
                  PropertyPaneTextField('addPageTitle', { label: 'Add Screen Title' }),
                  PropertyPaneTextField('editPageTitle', { label: 'Edit Screen Title' }),
                  PropertyPaneTextField('viewPageTitle', { label: 'View Screen Title' }),
                  PropertyPaneTextField('addSuccessMessage', { label: 'Success Message (Add)', placeholder: 'Default: Item created.' }),
                  PropertyPaneTextField('editSuccessMessage', { label: 'Success Message (Edit)', placeholder: 'Default: Item updated.' }),
                  PropertyPaneChoiceGroup('themeColor', {
                    label: 'Theme Color',
                    options: [
                      { key: 'siteTheme', text: 'Site Default', iconProps: { officeFabricIconFontName: 'Color' } },
                      { key: '#3b82f6', text: 'Blue', iconProps: { officeFabricIconFontName: 'Color' } },
                      { key: '#107c10', text: 'Green', iconProps: { officeFabricIconFontName: 'Color' } },
                      { key: '#d13438', text: 'Red', iconProps: { officeFabricIconFontName: 'Color' } },
                      { key: '#0078d4', text: 'SharePoint Blue', iconProps: { officeFabricIconFontName: 'Color' } },
                      { key: '#8764b8', text: 'Purple', iconProps: { officeFabricIconFontName: 'Color' } },
                      { key: '#ff8c00', text: 'Orange', iconProps: { officeFabricIconFontName: 'Color' } },
                      { key: '#008272', text: 'Teal', iconProps: { officeFabricIconFontName: 'Color' } }
                    ]
                  })
                ]
              }, (this as any).getConfigurationActionGroup()
            ]
          },
          // --- 2. FORM CONFIGS (From Helper) ---
          getGroup('add', 'Add Form Configuration'),
          getGroup('edit', 'Edit Form Configuration'),
          getGroup('view', 'View Form Configuration'),

          // --- 2. CHILD LIST LINKING (UPDATED) ---
          {
            header: { description: 'Child List Linking' },
            groups: [
              {
                groupName: 'Child List Setup',
                groupFields: [
                  PropertyPaneLabel('lblChildInfo', {
                    text: 'Link a child list (e.g. Cities) to this parent list.'
                  }),

                  // 1. TOGGLE BUTTON (Open/Close Config)
                  PropertyPaneCustomField({
                    key: 'childListConfigBtn',
                    onRender: (elem) => {
                      const isOpen = this.properties.isConfiguringChild;
                      const btn = React.createElement('button', {
                        style: {
                          width: '100%', padding: '10px', cursor: 'pointer',
                          background: isOpen ? '#f3f2f1' : '#eff6ff',
                          border: '1px solid #0078d4',
                          color: isOpen ? '#333' : '#0078d4',
                          fontWeight: 600
                        },
                        onClick: () => {
                          //  Use dot notation now that interface is updated
                          this.properties.isConfiguringChild = !this.properties.isConfiguringChild;
                          this.context.propertyPane.refresh();
                        }
                      }, isOpen ? "Close Configuration" : "Configure Child List");
                      ReactDom.render(btn, elem);
                    }
                  }),

                  // 2. CONFIGURATION FIELDS (Only show if open)
                  ...(this.properties.isConfiguringChild ? [
                    PropertyPaneTextField('childConfig_title', { label: 'Section Title', placeholder: 'e.g. Cities' }),
                    PropertyPaneDropdown('childConfig_list', { label: 'Select Child List', options: this.lists }),
                    PropertyPaneTextField('childConfig_ref', { label: 'Foreign Key Field (Internal Name)', description: 'The Lookup field in the Child list.' }),
                    PropertyPaneChoiceGroup('childConfig_mode', {
                      label: 'UI Mode',
                      options: [
                        { key: 'row', text: 'Row Based (Grid)', iconProps: { officeFabricIconFontName: 'Table' } },
                        { key: 'form', text: 'Form Based (Modal)', iconProps: { officeFabricIconFontName: 'OpenInNewWindow' } }
                      ]
                    }),
                    PropertyPaneTextField('childConfig_fields', { label: 'Visible Fields (Internal Names, comma separated)' }),

                    // SAVE BUTTON
                    PropertyPaneButton('btnSaveChild', {
                      text: 'Save Child Configuration',
                      buttonType: PropertyPaneButtonType.Primary,
                      icon: 'Save',
                      onClick: () => {
                        const newConfig: IChildListConfig = {
                          title: this.properties.childConfig_title || 'Child Items',
                          childListTitle: this.properties.childConfig_list || '',
                          foreignKeyField: this.properties.childConfig_ref || '',
                          uiMode: (this.properties.childConfig_mode as any) || 'row',
                          visibleFields: (this.properties.childConfig_fields || '').split(',').map((s: string) => s.trim())
                        };
                        this.properties.childConfigs = [newConfig];
                        this.properties.isConfiguringChild = false;
                        this.context.propertyPane.refresh();
                        void Swal.fire({ icon: 'success', title: 'Saved', text: 'Child list linked successfully.' });
                      }
                    }),

                    // CLEAR BUTTON (Removes the link)
                    PropertyPaneButton('btnClearChild', {
                      text: 'Remove Link',
                      buttonType: PropertyPaneButtonType.Normal,
                      icon: 'Delete',
                      onClick: () => {
                        this.properties.childConfigs = [];
                        this.properties.isConfiguringChild = false;
                        this.properties.childConfig_title = '';
                        this.properties.childConfig_list = '';
                        this.properties.childConfig_ref = '';
                        this.context.propertyPane.refresh();
                        void Swal.fire({ icon: 'success', title: 'Removed', text: 'Child list link has been removed.' });
                      }
                    })
                  ] : []),

                  // 3. STATUS LABEL
                  PropertyPaneLabel('currConfig', {
                    text: (this.properties.childConfigs && this.properties.childConfigs.length > 0)
                      ? `✅ Linked to: ${this.properties.childConfigs[0].childListTitle}`
                      : '❌ No child list linked.'
                  })
                ]
              },
              // 4. PANEL ACTIONS (Save/Cancel for the whole web part)
              (this as any).getConfigurationActionGroup()
            ]
          },
          // --- 3. SECTIONS & LAYOUT (From Hold) ---
          {
            header: { description: 'Form Layout & Sections' },
            groups: [
              {
                groupName: 'General Layout',
                groupFields: [
                  PropertyPaneChoiceGroup('formLayout', {
                    label: 'Columns per Row',
                    options: [
                      { key: 'single', text: 'One Column', iconProps: { officeFabricIconFontName: 'SingleColumn' } },
                      { key: 'double', text: 'Two Columns', iconProps: { officeFabricIconFontName: 'DoubleColumn' } }
                    ]
                  })
                ]
              },
              {
                groupName: 'Grouping',
                groupFields: [
                  PropertyPaneLabel('lblSections', {
                    text: 'Configure sections to group fields in Add/Edit/View forms. Select your fields in the previous tabs first.'
                  }),
                  PropertyPaneCustomField({
                    key: 'sectionEditorConfig',
                    onRender: (elem) => {
                      try {
                        if (this.properties.isConfiguringSections) {
                          const editor = React.createElement(SectionEditor, {
                            sections: this.properties.formSections || [],
                            availableFields: this.fields.filter(f => f.key !== 'ContentType'),
                            onSave: (newSections) => {
                              this.properties.formSections = newSections;
                              this.properties.isConfiguringSections = false;
                              this.context.propertyPane.refresh();
                              this.render();
                            },
                            onCancel: () => {
                              this.properties.isConfiguringSections = false;
                              this.context.propertyPane.refresh();
                            }
                          });
                          ReactDom.render(editor, elem);
                        } else {
                          const count = (this.properties.formSections || []).length;
                          const btn = React.createElement('button', {
                            onClick: () => {
                              this.properties.isConfiguringSections = true;
                              this.context.propertyPane.refresh();
                            },
                            style: {
                              marginTop: 15, width: '100%', padding: '10px',
                              backgroundColor: '#fff', border: '1px solid #0078d4',
                              color: '#0078d4', borderRadius: '4px', cursor: 'pointer', fontWeight: 600
                            }
                          }, `Configure Form Sections/Groups (${count})`);
                          ReactDom.render(btn, elem);
                        }
                      } catch (error: any) {
                        void LoggerService.log('PowerFormWebPart-renderSectionEditor', 'Medium', 'Config', error.message || JSON.stringify(error));
                      }
                    }
                  })
                ]
              },
              {
                groupName: 'Section Display Mode',
                groupFields: [
                  PropertyPaneChoiceGroup('sectionLayout', {
                    label: 'Display Sections As:',
                    options: [
                      { key: 'none', text: 'None (Plain List)', iconProps: { officeFabricIconFontName: 'BulletedList' } },
                      { key: 'stacked', text: 'Stacked (Current)', iconProps: { officeFabricIconFontName: 'Sections' } },
                      { key: 'tabs', text: 'Tabs', iconProps: { officeFabricIconFontName: 'TabCenter' } },
                      { key: 'wizard', text: 'Wizard (Step-by-Step)', iconProps: { officeFabricIconFontName: 'DoubleChevronRight' } }
                    ]
                  })
                ]
              }, (this as any).getConfigurationActionGroup()
            ]
          },
          // --- 4. LIST VIEW (From Helper) ---
          getGroup('list', 'List View Configuration'),
          // --- 5. MAINTENANCE & MIGRATION (From Hold) ---
          {
            header: { description: 'Maintenance & Migration' },
            groups: [
              {
                groupName: 'Configuration Management',
                groupFields: [
                  PropertyPaneLabel('lblExport', {
                    text: 'Export the current configuration to a JSON file. This file can be imported into another instance of this web part.'
                  }) as any,
                  PropertyPaneButton('btnExport', {
                    text: 'Export Configuration',
                    buttonType: PropertyPaneButtonType.Primary,
                    icon: 'Download',
                    onClick: this.exportConfiguration
                  }) as any
                ]
              },
              {
                groupName: 'System Logs On/off',
                groupFields: [
                  PropertyPaneToggle('enableLogging', {
                    label: 'Enable System Logging',
                    onText: 'On', offText: 'Off',
                    checked: this.properties.enableLogging,
                    disabled: !this.properties.isLogSetupConfirmed // Safety Check  
                  }),
                  PropertyPaneToggle('showUserAlerts', {
                    label: 'Show Errors to Users (Popup)',
                    onText: 'Show', offText: 'Hide',
                    checked: this.properties.showUserAlerts,
                    disabled: !this.properties.enableLogging
                  }),
                  PropertyPaneLabel('lblLogs', { text: 'View system errors and logs captured by this web part.' }) as any,
                  PropertyPaneButton('btnLogs', {
                    text: 'Open Log Viewer',
                    buttonType: PropertyPaneButtonType.Hero,
                    icon: 'ComplianceAudit',
                    onClick: () => { (this as any).showLogViewer = true; this.render(); }
                  }) as any
                ]
              },
              //  MERGED PERMISSION OVERRIDES (Requested in prompt) 
              {
                groupName: 'Global Permission Overrides',
                groupFields: [
                  PropertyPaneLabel('lblOverride', {
                    text: 'Check these boxes to force-hide buttons for ALL users (including Admins).'
                  }) as any,
                  PropertyPaneCheckbox('overrideAdd', { text: 'Hide "New" Button' }),
                  PropertyPaneCheckbox('overrideEdit', { text: 'Hide "Edit" Button' }),
                  PropertyPaneCheckbox('overrideDelete', { text: 'Hide "Delete" Button' }),
                  PropertyPaneToggle('isLargeList', {
                    label: "Force Large List Mode",
                    onText: "Enabled (Pagination On)",
                    offText: "Auto (Based on Count)"
                  })
                ]
              }, (this as any).getConfigurationActionGroup()
            ]
          },
          {
            header: { description: "Configure Notifications" },
            groups: [
              {
                groupName: "General Settings",
                groupFields: [
                  PropertyPaneToggle('enableNotification', { label: "Enable Module" }),
                  PropertyPaneLabel('lblNotifInfo', { text: "Ensure 'BAZnotifications' list exists." }),
                  PropertyPaneTextField('notifRoleName', { label: "Role Name (for Everyone)" }),
                  PropertyPaneButton('btnNotifPerms', {
                    text: "Configure Permissions",
                    buttonType: PropertyPaneButtonType.Primary,
                    onClick: this.configureNotificationPermissions.bind(this),
                    disabled: !this.properties.enableNotification
                  })
                ]
              },
              // ADD SECTION
              {
                groupName: "Add Notifications",
                groupFields: [
                  PropertyPaneToggle('enableNotifAdd', { label: "Notify on Add" }),
                  PropertyPaneTextField('msgNotifAdd', { label: "Default Message", multiline: true, disabled: !this.properties.enableNotifAdd }),
                  this.getGroupPickerControl('Default Groups (Empty = Everyone)', 'groupsNotifAdd'),
                  this.getCustomRuleEditorControl("Advanced Add Rules", "rulesAdd", "isConfiguringRulesAdd")
                ]
              },
              // UPDATE SECTION
              {
                groupName: "Update Notifications",
                groupFields: [
                  PropertyPaneToggle('enableNotifUpdate', { label: "Notify on Update" }),
                  PropertyPaneTextField('msgNotifUpdate', { label: "Default Message", multiline: true, disabled: !this.properties.enableNotifUpdate }),
                  this.getGroupPickerControl('Default Groups (Empty = Everyone)', 'groupsNotifUpdate'),
                  this.getCustomRuleEditorControl("Advanced Update Rules", "rulesUpdate", "isConfiguringRulesUpdate")
                ]
              },
              // DELETE SECTION
              {
                groupName: "Delete Notifications",
                groupFields: [
                  PropertyPaneToggle('enableNotifDelete', { label: "Notify on Delete" }),
                  PropertyPaneTextField('msgNotifDelete', { label: "Default Message", multiline: true, disabled: !this.properties.enableNotifDelete }),
                  this.getGroupPickerControl('Default Groups (Empty = Everyone)', 'groupsNotifDelete'),
                  this.getCustomRuleEditorControl("Advanced Delete Rules", "rulesDelete", "isConfiguringRulesDelete")
                ]
              },
              // VIEW SECTION
              {
                groupName: "View Notifications",
                groupFields: [
                  PropertyPaneToggle('enableNotifView', { label: "Notify on View" }),
                  PropertyPaneTextField('msgNotifView', { label: "Default Message", multiline: true, disabled: !this.properties.enableNotifView }),
                  this.getGroupPickerControl('Default Groups (Empty = Everyone)', 'groupsNotifView'),
                  this.getCustomRuleEditorControl("Advanced View Rules", "rulesView", "isConfiguringRulesView")
                ]
              },
              (this as any).getConfigurationActionGroup()
            ]
          }
        );
      }
      return { pages: pages };
    } catch (error: any) {
      void LoggerService.log('getPropertyPaneConfiguration', 'High', 'Config', error.message);
      return { pages: [] };
    }
  }
}