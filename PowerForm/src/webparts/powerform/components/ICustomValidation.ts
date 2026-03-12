/**
 * Defines when the validation rule should run.
 * - 'change': Real-time validation as the user types.
 * - 'blur': Validates when the field loses focus.
 * - 'submit': Final check before the form is sent to SharePoint.
 */
export type ValidationTrigger = 'change' | 'blur' | 'submit';

/**
 * Supported validation logic types for the PowerForm engine.
 */
export type ValidationType = 'regex' | 'range' | 'compare' | 'custom';

/**
 * Represents a single validation rule applied to a field.
 */
export interface ICustomValidationRule {
  /** * Unique ID for this rule (usually a timestamp or GUID).
   */
  id: string;

  /** * The event that triggers the validation.
   */
  trigger: ValidationTrigger;

  /**
   * The category of validation logic to apply.
   */
  type: ValidationType;

  /** * The Regular Expression string.
   * Required if type is 'regex'.
   */
  pattern?: string;

  /** Minimum numeric value or Date string. */
  min?: number | string;

  /** Maximum numeric value or Date string. */
  max?: number | string;

  // --- COMPARISON SETTINGS ---
  
  /** The Internal Name of the field to compare against (e.g., 'StartDate'). */
  otherField?: string;

  /** * The comparison operator.
   * eq (=), ne (!=), gt (>), lt (<), ge (>=), le (<=)
   */
  operator?: 'eq' | 'ne' | 'gt' | 'lt' | 'ge' | 'le';

  // --- CUSTOM SCRIPT SETTINGS ---

  /** * The body of the custom JavaScript function.
   * Expected to return a boolean or a Promise<boolean>.
   */
  fnBody?: string;

  /** * The localized error message shown to the user on failure.
   */
  message: string;
}

/**
 * A dictionary storing all validation rules for the Web Part.
 * Key: Field Internal Name (e.g., "Title", "Project_x0020_Status")
 * Value: Array of validation rules applied to that field.
 */
export interface IValidationConfig {
  [fieldInternalName: string]: ICustomValidationRule[];
}