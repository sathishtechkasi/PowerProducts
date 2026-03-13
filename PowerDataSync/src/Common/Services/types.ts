/**
 * Enum for supported SharePoint Column Types.
 * Using 'const enum' improves performance during the Heft build process.
 */
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
  Taxonomy = "TaxonomyFieldType",
  TaxonomyMulti = "TaxonomyFieldTypeMulti",
  Attachments = "Attachments"
}

/**
 * Represents the configuration for a SharePoint Site Column or List Field.
 */
export interface IColumnDefinition {
  /** Display name shown in the UI */
  title: string;
  /** Internal system name (e.g., 'My_x0020_Field') */
  internalName: string;
  /** The SharePoint data type */
  type: ColumnType;
  /** Unique ID (GUID) for the field */
  id: string; 
  /** Group name for Site Column organization */
  group: string;
  /** Whether the field is mandatory */
  required?: boolean;
  /** Character limit for Text/Note fields */
  maxLength?: number; 
  /** Prevents duplicate values in the column */
  enforceUnique?: boolean; 
  /** Options for Choice and MultiChoice types */
  choices?: string[]; 
  /** For Lookup fields: The Title or GUID of the target list */
  lookupListTitle?: string; 
  /** For Lookup fields: The internal name of the field to retrieve */
  lookupField?: string; 
  /** For Lookup fields: Optional additional field to display */
  showField?: string; 
  /** For User fields: 1 = People Only, 0 = People and Groups */
  userSelectionMode?: number; 
  /** Flag for Multi-value support (Choice, Lookup, User) */
  isMulti?: boolean;
  /** Optional description for the field */
  description?: string;
}

/**
 * Represents a SharePoint Content Type definition.
 */
export interface IContentTypeDefinition {
  /** Name of the Content Type */
  name: string;
  /** Detailed description */
  description: string;
  /** The full Content Type ID (e.g., '0x010100...') */
  id: string; 
  /** Group name for organization */
  group: string;
  /** Array of Site Column GUIDs linked to this Content Type */
  fieldGuids: string[]; 
}