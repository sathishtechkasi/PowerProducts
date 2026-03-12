import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { ICustomAction } from './ICustomAction';
import { LoggerService } from './LoggerService';

export interface ICustomActionEditorProps {
  label: string;
  actions: ICustomAction[];
  onSave: (actions: ICustomAction[]) => void;
  onCancel: () => void;
}

export interface ICustomActionEditorState {
  actions: ICustomAction[];
  editing: Partial<ICustomAction>;
}

export class CustomActionEditor extends React.Component<ICustomActionEditorProps, ICustomActionEditorState> {
  constructor(props: ICustomActionEditorProps) {
    super(props);
    try {
      this.state = {
        actions: props.actions ? [...props.actions] : [],
        editing: { id: '', title: '', url: '', icon: '', target: '_self', className: '' }
      };
    } catch (error:any) {
      void LoggerService.log('CustomActionEditor-constructor', 'High', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _saveAction = (): void => {
    try {
      if (!this.state.editing.title || !this.state.editing.url) return;
      
      const { editing, actions } = this.state;
      const isUpdate = actions.some(a => a.id === editing.id);
      
      const newAction = {
        ...editing,
        id: editing.id || `action_${Date.now()}`
      } as ICustomAction;

      const newActions = isUpdate 
        ? actions.map(a => a.id === newAction.id ? newAction : a)
        : [...actions, newAction];

      this.setState({ actions: newActions, editing: { id: '', title: '', url: '', icon: '', target: '_self', className: '' } });
    } catch (error:any) {
      void LoggerService.log('CustomActionEditor-saveAction', 'High', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _editAction = (action: ICustomAction): void => {
    try {
      this.setState({ editing: { ...action } });
    } catch (error:any) {
      void LoggerService.log('CustomActionEditor-editAction', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _deleteAction = (id: string): void => {
    try {
      this.setState(prev => ({
        actions: prev.actions.filter(a => a.id !== id),
        editing: prev.editing.id === id ? { id: '', title: '', url: '', icon: '', target: '_self', className: '' } : prev.editing
      }));
    } catch (error:any) {
      void LoggerService.log('CustomActionEditor-deleteAction', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _clearForm = (): void => {
    try {
      this.setState({ editing: { id: '', title: '', url: '', icon: '', target: '_self', className: '' } });
    } catch (error:any) {
      void LoggerService.log('CustomActionEditor-clearForm', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  public render(): React.ReactElement<ICustomActionEditorProps> {
    const { actions, editing } = this.state;
    const isEditingExisting = actions.some(a => a.id === editing.id);

    return (
      <div style={{ padding: '10px 0', fontSize: 13 }}>
        <h3 style={{ margin: '0 0 15px 0', color: '#323130' }}>{this.props.label}</h3>

        {/* --- LIST OF CONFIGURED ACTIONS --- */}
        {actions.length > 0 && (
          <div style={{ marginBottom: 20, border: '1px solid #edebe9', borderRadius: 2, overflow: 'hidden' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', background: '#fff' }}>
              <thead>
                <tr style={{ background: '#faf9f8', borderBottom: '1px solid #edebe9' }}>
                  <th style={{ padding: 8, fontWeight: 600 }}>Title</th>
                  <th style={{ padding: 8, fontWeight: 600 }}>URL</th>
                  <th style={{ padding: 8, fontWeight: 600 }}>Target</th>
                  <th style={{ padding: 8, width: 80 }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {actions.map(a => (
                  <tr key={a.id} style={{ borderBottom: '1px solid #f3f2f1', background: editing.id === a.id ? '#eff6ff' : 'transparent' }}>
                    <td style={{ padding: 8 }}>{a.icon && <i className={`ms-Icon ms-Icon--${a.icon}`} style={{ marginRight: 6 }} />} {a.title}</td>
                    <td style={{ padding: 8 }}>
                      <div style={{ maxWidth: 150, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={a.url}>{a.url}</div>
                    </td>
                    <td style={{ padding: 8 }}>{a.target === '_blank' ? 'New Tab' : 'Same Page'}</td>
                    <td style={{ padding: 8, whiteSpace: 'nowrap' }}>
                      <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editAction(a)} styles={{ root: { color: '#0078d4', height: 24, width: 24 } }} />
                      <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._deleteAction(a.id)} styles={{ root: { color: '#d13438', height: 24, width: 24 } }} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* --- ADD / EDIT FORM --- */}
        <div style={{ border: '1px solid #e1dfdd', padding: 15, borderRadius: 2, background: isEditingExisting ? '#fdfdfd' : '#faf9f8' }}>
          <h4 style={{ margin: '0 0 15px', color: isEditingExisting ? '#0078d4' : '#323130' }}>{isEditingExisting ? 'Edit Button' : 'Add New Button'}</h4>
          
          <div style={{ display: 'flex', gap: 15, marginBottom: 15 }}>
            <TextField 
              label="Button Title" 
              required 
              value={editing.title || ''} 
              onChange={(_, val) => this.setState({ editing: { ...editing, title: val || '' } })} 
              style={{ flex: 1 }} 
            />
            <TextField 
              label="Icon Name (Fluent UI)" 
              placeholder="e.g., Save, Mail, Print"
              value={editing.icon || ''} 
              onChange={(_, val) => this.setState({ editing: { ...editing, icon: val || '' } })} 
              style={{ flex: 1 }} 
            />
          </div>

          <div style={{ marginBottom: 15 }}>
            <TextField 
              label="Action URL or JavaScript" 
              required 
              multiline rows={2}
              placeholder="https://... OR javascript:alert('hello');"
              value={editing.url || ''} 
              onChange={(_, val) => this.setState({ editing: { ...editing, url: val || '' } })} 
            />
          </div>

          <div style={{ display: 'flex', gap: 15, marginBottom: 15 }}>
            <div style={{ flex: 1 }}>
              <Dropdown 
                label="Open Link In"
                options={[ { key: '_self', text: 'Same Page' }, { key: '_blank', text: 'New Tab' } ]}
                selectedKey={editing.target || '_self'}
                onChange={(_, opt) => this.setState({ editing: { ...editing, target: (opt?.key as any) || '_self' } })}
              />
              <div style={{ fontSize: 11, color: '#605e5c', marginTop: 4 }}>
                *JavaScript actions ignore this setting.
              </div>
            </div>

            <TextField 
              label="Custom CSS Class (Optional)" 
              placeholder="e.g., ms-Button--primary"
              value={editing.className || ''}
              onChange={(_, val) => this.setState({ editing: { ...editing, className: val || '' } })}
            />
          </div>

          <div style={{ display: 'flex', gap: 8 }}>
            <PrimaryButton 
              text={isEditingExisting ? 'Update Button' : 'Add Button'} 
              disabled={!editing.title || !editing.url} 
              onClick={this._saveAction}
            />
            {isEditingExisting && (
              <DefaultButton text="Cancel Edit" onClick={this._clearForm} />
            )}
          </div>
        </div>

        {/* --- FOOTER BUTTONS (Global Save) --- */}
        <div style={{ marginTop: 20, paddingTop: 15, borderTop: '1px solid #edebe9', display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
           <PrimaryButton 
            text="Save & Close" 
            onClick={() => {
              try {
                this.props.onSave(this.state.actions);
              } catch (error:any) {
                void LoggerService.log('CustomActionEditor-onSave', 'High', 'Config', error instanceof Error ? error.message : String(error));
              }
            }} 
           />
           <DefaultButton text="Cancel" onClick={this.props.onCancel} />
        </div>
      </div>
    );
  }
}