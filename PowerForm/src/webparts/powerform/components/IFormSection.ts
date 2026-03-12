/**
 * Represents a logical section/group within the PowerForm.
 * Used for organizing fields into tabs, wizard steps, or collapsible groups.
 */
export interface IFormSection {
  /** * Unique identifier for the section. 
   * Recommended: Use a GUID or stringified timestamp.
   */
  id: string;

  /** * The display title of the section shown in the UI.
   */
  title: string;

  /** * Array of SharePoint InternalNames for fields belonging to this section.
   */
  fields: string[];

  /** * Determines the display order of the section (lower numbers first).
   */
  order: number;

  /**
   * Optional: If true, the section starts in a collapsed state.
   * Useful for "Advanced" or "Metadata" sections.
   */
  isCollapsible?: boolean;

  /**
   * Optional: Column layout for this specific section.
   * Allows overriding the global web part layout.
   */
  columns?: 1 | 2 | 3;

  /**
   * Optional: Description or help text shown below the section title.
   */
  description?: string;
}