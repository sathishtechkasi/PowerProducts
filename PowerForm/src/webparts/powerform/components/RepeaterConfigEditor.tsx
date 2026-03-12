import * as React from 'react';
import { IRepeaterColumn } from './IPowerFormProps';
import { IconButton, DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { LoggerService } from './LoggerService';

const styles: { [key: string]: React.CSSProperties } = {
  row: { background: '#faf9f8', padding: '12px', marginBottom: '12px', border: '1px solid #e1dfdd', borderRadius: '2px' },
  subPanel: { marginTop: '10px', padding: '10px', background: '#f3f2f1', borderRadius: '2px' }
};

export interface IRepeaterConfigEditorProps {
  fieldKey: string;
  currentConfig: IRepeaterColumn[];
  onSave: (config: IRepeaterColumn[]) => void;
  onCancel: () => void;
  onClear: () => void;
}

export class RepeaterConfigEditor extends React.Component<IRepeaterConfigEditorProps, { columns: IRepeaterColumn[] }> {
  constructor(props: IRepeaterConfigEditorProps) {
    super(props);
    try {
      this.state = { columns: props.currentConfig ? [...props.currentConfig] : [] };
    } catch (error:any) {
      void LoggerService.log('RepeaterConfigEditor-constructor', 'High', 'Config', error instanceof Error ? error.message : String(error));
      this.state = { columns: [] };
    }
  }

  private _addCol = (): void => {
    try {
      this.setState(prev => ({ columns: [...prev.columns, { key: `col_${Date.now()}`, name: '', type: 'text' }] }));
    } catch (error:any) {
      void LoggerService.log('RepeaterConfigEditor-addCol', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _removeCol = (idx: number): void => {
    try {
      const newCols = [...this.state.columns];
      newCols.splice(idx, 1);
      this.setState({ columns: newCols });
    } catch (error:any) {
      void LoggerService.log('RepeaterConfigEditor-removeCol', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _updateCol = (idx: number, prop: keyof IRepeaterColumn, val: any): void => {
    try {
      const newCols = [...this.state.columns];
      newCols[idx] = { ...newCols[idx], [prop]: val };
      this.setState({ columns: newCols });
    } catch (error:any) {
      void LoggerService.log('RepeaterConfigEditor-updateCol', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  public render(): React.ReactElement<IRepeaterConfigEditorProps> {
    try {
      return (
        <div style={{ padding: '10px 0' }}>
          <h3 style={{ margin: '0 0 15px 0', fontSize: '15px', color: '#323130' }}>Repeater Grid Configuration</h3>
          <p style={{ fontSize: '12px', color: '#605e5c', marginBottom: 20 }}>
            Define the columns that will appear inside this data grid. Data is stored as JSON text in the underlying SharePoint multiline field.
          </p>

          <div style={{ maxHeight: '400px', overflowY: 'auto', paddingRight: '5px' }}>
            {this.state.columns.map((c, idx) => (
              <div key={c.key} style={styles.row}>
                <div style={{ display: 'flex', gap: 15, alignItems: 'flex-start' }}>
                  <div style={{ flex: 1 }}>
                    <TextField label="Column Name" required value={c.name} onChange={(_, val) => this._updateCol(idx, 'name', val || '')} />
                  </div>
                  <div style={{ flex: 1 }}>
                    <Dropdown 
                      label="Input Type" 
                      options={[
                        { key: 'text', text: 'Single Line of Text' },
                        { key: 'number', text: 'Number' },
                        { key: 'date', text: 'Date' },
                        { key: 'choice', text: 'Dropdown (Single)' },
                        { key: 'multichoice', text: 'Dropdown (Multi)' }
                      ]} 
                      selectedKey={c.type || 'text'} 
                      onChange={(_, opt) => this._updateCol(idx, 'type', opt?.key)} 
                    />
                  </div>
                  <IconButton iconProps={{ iconName: 'Delete' }} title="Remove Column" onClick={() => this._removeCol(idx)} styles={{ root: { color: '#d13438', marginTop: 28 }, rootHovered: { background: '#fde7e9' } }} />
                </div>

                {(c.type === 'choice' || c.type === 'multichoice') && (
                  <div style={styles.subPanel}>
                    <TextField 
                      label="Dropdown Options" 
                      multiline rows={2} 
                      placeholder="Option 1, Option 2, Option 3 (comma separated)" 
                      value={c.options || ''} 
                      onChange={(_, val) => this._updateCol(idx, 'options', val || '')} 
                    />
                  </div>
                )}

                {c.type === 'date' && (
                  <div style={styles.subPanel}>
                    <div style={{ display: 'flex', gap: 15 }}>
                      <Dropdown 
                        label="Date Validation Rule" 
                        options={[
                          { key: '', text: 'None' },
                          { key: 'future', text: 'Must be in the Future' },
                          { key: 'past', text: 'Must be in the Past' },
                          { key: 'future_n', text: 'Older than N days' },
                          { key: 'past_n', text: 'Due in N days' }
                        ]} 
                        selectedKey={c.dateRule || ''} 
                        onChange={(_, opt) => this._updateCol(idx, 'dateRule', opt?.key)} 
                        style={{ flex: 1 }}
                      />
                      {(c.dateRule === 'future_n' || c.dateRule === 'past_n') && (
                        <TextField 
                          label="Number of Days (N)" 
                          type="number" 
                          value={c.dateDays?.toString() || ''} 
                          onChange={(_, val) => this._updateCol(idx, 'dateDays', parseInt(val || '0', 10))} 
                          style={{ flex: 1 }}
                        />
                      )}
                    </div>
                  </div>
                )}

                <div style={{ display: 'flex', gap: 20, marginTop: 15 }}>
                  <Checkbox label="Required Field" checked={!!c.required} onChange={(_, checked) => this._updateCol(idx, 'required', checked)} />
                  <Checkbox label="Unique Value" checked={!!c.unique} onChange={(_, checked) => this._updateCol(idx, 'unique', checked)} />
                </div>
              </div>
            ))}
          </div>

          <PrimaryButton text="+ Add Column" onClick={this._addCol} style={{ marginTop: 15, width: '100%', background: '#fff', color: '#0078d4', border: '1px dashed #0078d4' }} />

          <div style={{ marginTop: 20, display: 'flex', gap: 10, borderTop: '1px solid #edebe9', paddingTop: 15 }}>
            <PrimaryButton text="Save Grid Config" onClick={() => this.props.onSave(this.state.columns)} style={{ flex: 1 }} />
            <DefaultButton text="Cancel" onClick={this.props.onCancel} style={{ flex: 1 }} />
          </div>

          <div style={{ marginTop: 15, textAlign: 'center' }}>
            <DefaultButton 
              text="Remove Configuration" 
              onClick={this.props.onClear} 
              styles={{ root: { width: '100%', color: '#d13438', borderColor: '#d13438' }, rootHovered: { background: '#fde7e9' } }}
              title="Removes repeater configuration and reverts to standard text field"
            />
          </div>
        </div>
      );
    } catch (error:any) {
      void LoggerService.log('RepeaterConfigEditor-render', 'High', 'Config', error instanceof Error ? error.message : String(error));
      return <div>Error loading editor</div>;
    }
  }
}