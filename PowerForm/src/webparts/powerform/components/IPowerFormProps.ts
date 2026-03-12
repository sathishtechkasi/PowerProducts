import { SPHttpClient } from '@microsoft/sp-http';
import { CommonService } from '../../../Common/Services/CommonService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IValidationConfig } from './ICustomValidation';
import { ICustomAction } from './ICustomAction';
import { IViewConfig } from './ViewEditor';
import { IFormSection } from './IFormSection';

/**
 * Mapping between source and target fields for auto-population logic.
 */
export interface IColumnMapping {
  source: string;
  target: string;
}

/**
 * Configuration for type-ahead functionality in SharePoint forms.
 */
export interface IAutocompleteConfig {
  [fieldKey: string]: {
    sourceList: string;
    sourceField: string;
    sourceQuery?: string;
    additionalFields?: string[];
    columnMapping?: IColumnMapping[];
  };
}

/**
 * Enhanced lookup display settings for Fluent UI 8 dropdowns.
 */
export interface ILookupDisplayConfig {
  [fieldKey: string]: {
    additionalFields: string[];
    filterQuery?: string;
    columnMapping?: IColumnMapping[];
  };
}

/**
 * Parent-Child relationship configuration for filtered dropdowns.
 */
export interface ICascadeConfig {
  [childField: string]: {
    parentField: string;
    foreignKey: string;
    filterQuery?: string;
    additionalFields?: string[];
    columnMapping?: IColumnMapping[];
  };
}

/**
 * Configuration for Child/Sub-list relationships.
 */
export interface IChildListConfig {
  childListTitle: string;
  foreignKeyField: string;
  uiMode: 'row' | 'form';
  visibleFields: string[];
  title: string;
}

/**
 * Configuration for the JSON-based Repeater grid.
 */
export interface IRepeaterConfig {
  [fieldInternalName: string]: IRepeaterColumn[];
}

export interface IRepeaterColumn {
  key: string;
  name: string;
  type: 'text' | 'date' | 'choice' | 'multichoice' | 'number';
  options?: string;
  required?: boolean;
  width?: number;
  unique?: boolean;
  dateRule?: 'future' | 'past' | 'future_n' | 'past_n';
  dateDays?: number;
}

/**
 * Logic-based notification rules for specific item conditions.
 */
export interface INotificationRule {
  enabled: boolean;
  condition: string;
  message: string;
  targetGroups: string; 
  uniqueId: string;
}

/**
 * Primary props for the PowerForm component.
 */
export interface IPowerFormProps {
  // --- CORE CONTEXT ---
  selectedList: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  service: CommonService;
  context: WebPartContext;
  themeColor: string;

  // --- FORM DISPLAY CONFIGURATIONS ---
  addVisibleFields: string[];
  addFieldOrder: Record<string, number>;
  addReadOnlyFields: string[];
  addCustomScript?: string;
  addCustomStyle?: string;

  editVisibleFields: string[];
  editFieldOrder: Record<string, number>;
  editReadOnlyFields: string[];
  editCustomScript?: string;
  editCustomStyle?: string;

  viewVisibleFields: string[];
  viewFieldOrder: Record<string, number>;
  viewCustomScript?: string;
  viewCustomStyle?: string;

  // --- LIST VIEW SETTINGS ---
  isLargeList: boolean;
  listVisibleFields: string[];
  listFieldOrderMap: Record<string, number>;
  listFilterMap: Record<string, boolean>;
  listCustomScript?: string;
  listCustomStyle?: string;

  // --- LAYOUT & NAVIGATION ---
  formLayout?: 'single' | 'double';
  initialMode?: 'add' | 'edit' | 'view' | 'list';
  sectionLayout: 'none' | 'stacked' | 'tabs' | 'wizard';
  formSections?: IFormSection[];

  // --- TITLES & MESSAGES ---
  listPageTitle?: string;
  addPageTitle?: string;
  editPageTitle?: string;
  viewPageTitle?: string;
  addSuccessMessage?: string;
  editSuccessMessage?: string;

  // --- LOGIC & VALIDATION ---
  validationConfig?: IValidationConfig;
  autocompleteConfig?: IAutocompleteConfig;
  lookupDisplayConfig?: ILookupDisplayConfig;
  cascadeConfig?: ICascadeConfig;
  repeaterConfig?: IRepeaterConfig;

  // --- ACTIONS & PERMISSIONS ---
  editCustomActions?: ICustomAction[];
  viewCustomActions?: ICustomAction[];
  fieldPermissionConfig: Record<string, string[]>;
  views: IViewConfig[];
  defaultViewAllowedGroups: string[];
  overrideAdd?: boolean;
  overrideEdit?: boolean;
  overrideDelete?: boolean;

  // --- SYSTEM & LOGGING ---
  enableLogging: boolean;
  showUserAlerts: boolean;
  logListTitle: string;
  logRoleName: string;
  isLogInitialized: boolean;
  listGroupByField?: string;
  childConfigs?: IChildListConfig[];

  // --- NOTIFICATION ENGINE ---
  enableNotification: boolean;
  enableNotifAdd: boolean;
  msgNotifAdd: string;
  groupsNotifAdd: number[]; // Upgraded to strictly number IDs
  rulesAdd: INotificationRule[];

  enableNotifUpdate: boolean;
  msgNotifUpdate: string;
  groupsNotifUpdate: number[];
  rulesUpdate: INotificationRule[];

  enableNotifDelete: boolean;
  msgNotifDelete: string;
  groupsNotifDelete: number[];
  rulesDelete: INotificationRule[];

  enableNotifView: boolean;
  msgNotifView: string;
  groupsNotifView: number[];
  rulesView: INotificationRule[];
}
