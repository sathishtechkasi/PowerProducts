import { IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { IViewConfig } from './ViewEditor';
import { IChildListConfig } from './IPowerFormProps';

export const enum ColumnType {
  Text = "Text",
  Note = "Note",
  Number = "Number",
  Currency = "Currency",
  DateTime = "DateTime",
  Boolean = "Boolean",
  Choice = "Choice",
  MultiChoice = "MultiChoice",
  Lookup = "Lookup",
  LookupMulti = "LookupMulti",
  User = "User",
  UserMulti = "UserMulti",
  Url = "URL",
  Attachments = "Attachments"
}

export interface IExtendedPersonaProps extends IPersonaProps {
  key: string;
  text?: string;
  secondaryText?: string;
}

export interface ColumnDefinition {
  InternalName: string;
  EntityPropertyName: string;
  Title: string;
  TypeAsString: ColumnType | string;
  AllowMultipleValues?: boolean;
  Choices?: string[];
  LookupList?: string;
  Required?: boolean;
  FieldTypeKind?: number;
  MinimumValue?: number;
  MaximumValue?: number;
  EnforceUniqueValues?: boolean;
  LookupField?: string;
  DisplayFormat?: number;
}

export interface ListItem {
  Id: number;
  Attachments: boolean;
  [key: string]: any;
}

export interface IAttachmentInfo {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface ILookupOption extends IDropdownOption {
  itemData?: any;
}

export interface IPowerFormState {
  fields: ColumnDefinition[];
  items: ListItem[];
  formData: Record<string, any>;
  
  mode: 'list' | 'add' | 'edit' | 'view';
  itemId: number | undefined;

  loading: boolean;
  message: string;
  isPanelOpen: boolean;
  panelUrl: string;
  panelTitle: string;

  lookupOptions: Record<string, ILookupOption[]>;
  autocompleteOptions: Record<string, IComboBoxOption[]>;
  peopleOptions: Record<number, IExtendedPersonaProps>;
  pickerSearch: Record<string, string>;
  activePickerKey: string | null;

  attachmentsNew: File[];
  attachmentsDelete: string[];
  existingAttachments: IAttachmentInfo[];
  attachments: any[];

  formErrors: Record<string, string>;

  page: number;
  pageSize: number;
  totalItems: number;
  selectedItems: number[];
  filters: Record<string, { operator: string; value: string }>;
  searchText: string;
  sortField: string | null;
  sortDirection: 'asc' | 'desc';
  nextPageUrl?: string;
  
  pageCache: Record<number, {
    items: any[];
    nextHref?: string | null;
  }>;

  currentViewId: string;
  availableViews: IViewConfig[];
  activeViewFields: string[] | null;
  canSeeDefaultView: boolean;

  canAdd: boolean;
  canEdit: boolean;
  canView: boolean;
  canDelete: boolean;
  listId?: string;
  listUrl?: string;
  enableVersioning?: boolean;
  currentUserGroups: string[];
  urlReadOnlyFields: string[];
  activeSectionIndex: number;
  
  isBulkEditOpen: boolean;
  bulkEditField: string;
  bulkEditValue: string[] | null;
  isSaveDisabled: boolean;
  currentUser: any | null;
  
  childItems: Record<string, any[]>;
  isChildPanelOpen: boolean;
  activeChildConfig: IChildListConfig | null;
  activeChildItemIndex: number; 
  childFieldsCache: Record<string, ColumnDefinition[]>;
}