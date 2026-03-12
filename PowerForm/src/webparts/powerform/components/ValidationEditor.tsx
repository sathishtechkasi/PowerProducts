import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import {
  ICustomValidationRule,
  ValidationTrigger,
  ValidationType
} from './ICustomValidation';
import { LoggerService } from './LoggerService';

/* ───────────────────────────────────────────────────────────── */
export interface IValidationEditorProps {
  fieldKey: string;
  rules: ICustomValidationRule[];
  onSave: (updated: ICustomValidationRule[]) => void;
  onCancel: () => void;
}

export interface IValidationEditorState {
  rules: ICustomValidationRule[];
  editing: Partial<ICustomValidationRule & { id: string }>;
}

/* ───────────────────────────────────────────────────────────── */
const emptyRule = (): ICustomValidationRule => ({
  id: `val_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
  trigger: 'blur',
  type: 'regex',
  pattern: '',
  message: '',
  otherField: '',
  operator: 'eq'
});

/* ───────────────────────────────────────────────────────────── */
export class ValidationEditor extends React.Component<IValidationEditorProps, IValidationEditorState> {
  constructor(props: IValidationEditorProps) {
    super(props);
    try {
      this.state = {
        rules: props.rules ? [...props.rules] : [],
        editing: emptyRule()
      };
    } catch (error:any) {
      void LoggerService.log('ValidationEditor-constructor', 'High', 'Config', error instanceof Error ? error.message : String(error));
      this.state = { rules: [], editing: emptyRule() };
    }
  }

  /* ====== Helpers ============================================ */
  
  private _setEditing = <K extends keyof ICustomValidationRule>(k: K, v: any): void => {
    try {
      this.setState(prev => ({
        editing: { ...prev.editing, [k]: v }
      }));
    } catch (error:any) {
      void LoggerService.log('ValidationEditor-setEditing', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _canSave = (): boolean => {
    try {
      const r = this.state.editing;
      if (!r.message || r.message.trim() === '') return false;
      
      switch (r.type) {
        case 'regex':
          return !!r.pattern;
        case 'range':
          return r.min != null || r.max != null;
        case 'compare':
          return !!r.otherField && !!r.operator;
        case 'custom':
          return !!r.fnBody;
        default:
          return false;
      }
    } catch (error:any) {
      void LoggerService.log('ValidationEditor-canSave', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
      return false;
    }
  }

  private _saveRule = (): void => {
    try {
      if (!this._canSave()) return;
      const { editing, rules } = this.state;
      
      const isUpdate = rules.some(r => r.id === editing.id);
      if (isUpdate) {
        this.setState(prev => ({
          rules: prev.rules.map(r => r.id === editing.id ? (editing as ICustomValidationRule) : r),
          editing: emptyRule()
        }));
      } else {
        this.setState(prev => ({
          rules: [...prev.rules, editing as ICustomValidationRule],
          editing: emptyRule()
        }));
      }
    } catch (error:any) {
      void LoggerService.log('ValidationEditor-saveRule', 'High', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _editRule = (rule: ICustomValidationRule): void => {
    try {
      this.setState({ editing: { ...rule } });
    } catch (error:any) {
      void LoggerService.log('ValidationEditor-editRule', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _deleteRule = (id: string): void => {
    try {
      this.setState(prev => ({
        rules: prev.rules.filter(r => r.id !== id),
        editing: prev.editing.id === id ? emptyRule() : prev.editing
      }));
    } catch (error:any) {
      void LoggerService.log('ValidationEditor-deleteRule', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _cancelEdit = (): void => {
    try {
      this.setState({ editing: emptyRule() });
    } catch (error:any) {
      void LoggerService.log('ValidationEditor-cancelEdit', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  /* ====== Render Helpers ==================================== */
  
  private _renderRuleRow(r: ICustomValidationRule, isEditing: boolean): React.ReactNode {
    const desc = (() => {
      try {
        switch (r.type) {
          case 'regex': return `pattern: /${r.pattern}/`;
          case 'range': return `min: ${r.min ?? '–'} max: ${r.max ?? '–'}`;
          case 'compare': return `${r.operator} [${r.otherField}]`;
          case 'custom': return 'custom JS';
          default: return '-';
        }
      } catch {
        return 'Error';
      }
    })();

    return (
      <tr key={r.id} style={{ background: isEditing ? '#eff6ff' : 'transparent', borderBottom: '1px solid #f3f2f1' }}>
        <td style={{ padding: '8px', fontSize: '12px' }}>{r.trigger}</td>
        <td style={{ padding: '8px', fontSize: '12px' }}>{r.type}</td>
        <td style={{ padding: '8px', fontSize: '12px', color: '#605e5c' }}>{desc}</td>
        <td style={{ padding: '8px', fontSize: '12px' }}>{r.message}</td>
        <td style={{ padding: '4px', textAlign: 'right', whiteSpace: 'nowrap' }}>
          <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this._editRule(r)} styles={{ root: { color: '#0078d4', height: 24, width: 24 } }} />
          <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => this._deleteRule(r.id)} styles={{ root: { color: '#d13438', height: 24, width: 24 }, rootHovered: { background: '#fde7e9' } }} />
        </td>
      </tr>
    );
  }

  /* ====== Main Render ======================================= */
  public render(): React.ReactElement<IValidationEditorProps> {
    const { editing, rules } = this.state;
    const isEditingExisting = rules.some(r => r.id === editing.id);

    return (
      <div style={{ fontSize: 13, padding: '10px 0' }}>
        <h3 style={{ margin: '0 0 15px 0', fontSize: '14px', color: '#323130' }}>
          Validation for “{this.props.fieldKey}”
        </h3>

        {/* --- EXISTING RULES TABLE --- */}
        {rules.length > 0 && (
          <div style={{ overflowX: 'auto', marginBottom: 15, border: '1px solid #edebe9', borderRadius: 2, background: '#ffffff' }}>
            <table style={{ width: '100%', minWidth: '400px', borderCollapse: 'collapse' }}>
              <thead>
                <tr style={{ textAlign: 'left', borderBottom: '1px solid #edebe9', background: '#faf9f8' }}>
                  <th style={{ padding: '8px', fontWeight: 600, fontSize: '12px' }}>Trigger</th>
                  <th style={{ padding: '8px', fontWeight: 600, fontSize: '12px' }}>Type</th>
                  <th style={{ padding: '8px', fontWeight: 600, fontSize: '12px' }}>Config</th>
                  <th style={{ padding: '8px', fontWeight: 600, fontSize: '12px' }}>Message</th>
                  <th style={{ padding: '8px', width: 60 }} />
                </tr>
              </thead>
              <tbody>
                {rules.map(r => this._renderRuleRow(r, r.id === editing.id))}
              </tbody>
            </table>
          </div>
        )}

        {/* --- ADD / EDIT FORM --- */}
        <div style={{ border: '1px solid #e1dfdd', padding: 15, borderRadius: 2, background: isEditingExisting ? '#fdfdfd' : '#faf9f8' }}>
          <h4 style={{ marginTop: 0, marginBottom: 15, color: isEditingExisting ? '#0078d4' : '#323130' }}>
            {isEditingExisting ? 'Edit Rule' : 'Add Rule'}
          </h4>

          <div style={{ display: 'flex', gap: 15, marginBottom: 10 }}>
            <div style={{ flex: 1 }}>
              <Dropdown
                label="Trigger"
                options={[
                  { key: 'change', text: 'On Change (Immediate)' },
                  { key: 'blur', text: 'On Blur (Exit Field)' },
                  { key: 'submit', text: 'On Submit (Save)' }
                ]}
                selectedKey={editing.trigger}
                onChange={(_, opt) => this._setEditing('trigger', opt?.key as ValidationTrigger)}
              />
            </div>
            <div style={{ flex: 1 }}>
              <Dropdown
                label="Validation Type"
                options={[
                  { key: 'regex', text: 'Regex Pattern' },
                  { key: 'range', text: 'Numeric/Date Range' },
                  { key: 'compare', text: 'Compare Fields' },
                  { key: 'custom', text: 'Custom JS Script' }
                ]}
                selectedKey={editing.type}
                onChange={(_, opt) => this._setEditing('type', opt?.key as ValidationType)}
              />
            </div>
          </div>

          {/* DYNAMIC INPUTS BASED ON TYPE */}
          <div style={{ marginBottom: 15 }}>
            {editing.type === 'regex' && (
              <TextField 
                label="Pattern" 
                placeholder="^\d+$" 
                value={editing.pattern || ''} 
                onChange={(_, val) => this._setEditing('pattern', val || '')} 
              />
            )}

            {editing.type === 'range' && (
              <div style={{ display: 'flex', gap: 15 }}>
                <TextField label="Min Value" type="number" value={String(editing.min ?? '')} onChange={(_, val) => this._setEditing('min', val ? Number(val) : undefined)} style={{ flex: 1 }} />
                <TextField label="Max Value" type="number" value={String(editing.max ?? '')} onChange={(_, val) => this._setEditing('max', val ? Number(val) : undefined)} style={{ flex: 1 }} />
              </div>
            )}

            {editing.type === 'compare' && (
              <div style={{ display: 'flex', gap: 15 }}>
                <div style={{ flex: 1 }}>
                  <Dropdown
                    label="Operator"
                    options={[
                      { key: 'eq', text: 'Equal to (=)' },
                      { key: 'ne', text: 'Not Equal (!=)' },
                      { key: 'gt', text: 'Greater Than (>)' },
                      { key: 'ge', text: 'Greater/Equal (>=)' },
                      { key: 'lt', text: 'Less Than (<)' },
                      { key: 'le', text: 'Less/Equal (<=)' }
                    ]}
                    selectedKey={editing.operator || 'eq'}
                    onChange={(_, opt) => this._setEditing('operator', opt?.key)}
                  />
                </div>
                <div style={{ flex: 2 }}>
                  <TextField label="Other Field (Internal Name)" placeholder="e.g. StartDate" value={editing.otherField || ''} onChange={(_, val) => this._setEditing('otherField', val || '')} />
                </div>
              </div>
            )}

            {editing.type === 'custom' && (
              <TextField 
                label="JS Function Body (value => boolean)" 
                multiline rows={3} 
                placeholder="return value && value.length > 5;" 
                value={editing.fnBody || ''} 
                onChange={(_, val) => this._setEditing('fnBody', val || '')} 
              />
            )}
          </div>

          <TextField 
            label="Error Message" 
            placeholder="Please enter a valid value." 
            required 
            value={editing.message || ''} 
            onChange={(_, val) => this._setEditing('message', val || '')} 
            styles={{ root: { marginBottom: 15 } }}
          />

          <div style={{ display: 'flex', gap: 8 }}>
            <PrimaryButton text={isEditingExisting ? 'Update Rule' : 'Add Rule'} disabled={!this._canSave()} onClick={this._saveRule} />
            <DefaultButton text="Cancel Edit" onClick={this._cancelEdit} />
          </div>
        </div>

        {/* --- MAIN FOOTER BUTTONS --- */}
        <div style={{ marginTop: 20, paddingTop: 15, borderTop: '1px solid #edebe9', display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
          <PrimaryButton 
            text="Save & Close" 
            onClick={() => {
              try {
                this.props.onSave(this.state.rules);
              } catch (error:any) {
                void LoggerService.log('ValidationEditor-onSave-props', 'High', 'Config', error instanceof Error ? error.message : String(error));
              }
            }} 
          />
          <DefaultButton text="Cancel" onClick={this.props.onCancel} />
        </div>
      </div>
    );
  }
}