/**
 * Represents a custom button or action configured by the user 
 * in the Property Pane for the PowerForm Web Part.
 */
export interface ICustomAction {
  /** * Unique identifier for the action. 
   * Recommended: Use a GUID or stringified timestamp.
   */
  id: string;

  /** * The text label displayed on the button. 
   * Example: "Print Invoice", "Approve", "Go to Parent".
   */
  title: string;

  /** * The destination URL or JavaScript protocol link.
   * Supports dynamic placeholders: {ItemId}, {ListId}, {SiteUrl}, {WebUrl}.
   */
  url: string;

  /** * Optional Fluent UI (formerly Office UI Fabric) icon name.
   * Example: "Print", "OpenInNewWindow", "Share".
   * @see https://developer.microsoft.com/en-us/fluentui#/styles/web/icons
   */
  icon?: string;

  /**
   * Defines how the link should open.
   * - '_self': Same frame (standard for JS links)
   * - '_blank': New tab/window
   * - 'panel': (Advanced) Open the URL inside a SharePoint Side Panel
   */
  target?: '_self' | '_blank' | 'panel';

  /**
   * Optional CSS class for custom styling (e.g., making a button red for 'Delete').
   */
  className?: string;
}