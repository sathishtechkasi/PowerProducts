import * as React from 'react';
import { IRepeaterColumn } from './IPowerFormProps';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { LoggerService } from './LoggerService';

const styles = {
  input: { width: '100%', padding: '6px', border: '1px solid #c8c6c4', borderRadius: '2px', boxSizing: 'border-box' as const },
  select: { width: '100%', padding: '6px', border: '1px solid #c8c6c4', borderRadius: '2px' },
  errorBorder: { border: '1px solid #a4262c' },
  errorText: { color: '#a4262c', fontSize: '11px', marginTop: '4px', display: 'block' }
};

export interface IRepeaterInputProps {
  columns: IRepeaterColumn[];
  value: string;
  onChange: (newValue: string) => void;
  mode: 'edit' | 'view';
}

export class RepeaterInput extends React.Component<IRepeaterInputProps, { rows: any[] }> {
  constructor(props: IRepeaterInputProps) {
    super(props);
    this.state = { rows: this._parseValue(props.value) };
  }

  public componentDidUpdate(prevProps: IRepeaterInputProps): void {
    try {
      if (prevProps.value !== this.props.value) {
        const newRows = this._parseValue(this.props.value);
        if (JSON.stringify(newRows) !== JSON.stringify(this.state.rows)) {
          this.setState({ rows: newRows });
        }
      }
    } catch (error:any) {
      void LoggerService.log('RepeaterInput-componentDidUpdate', 'Medium', 'UI', error instanceof Error ? error.message : String(error));
    }
  }

  private _parseValue(val: string): any[] {
    try {
      if (!val) return [];
      const parsed = JSON.parse(val);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  private _updateRow = (index: number, colKey: string, val: any): void => {
    try {
      const newRows = [...this.state.rows];
      newRows[index] = { ...newRows[index], [colKey]: val };
      this.setState({ rows: newRows }, () => this.props.onChange(JSON.stringify(newRows)));
    } catch (error:any) {
      void LoggerService.log('RepeaterInput-updateRow', 'High', 'UserAction', error instanceof Error ? error.message : String(error));
    }
  }

  private _addRow = (): void => {
    try {
      const newRows = [...this.state.rows, {}];
      this.setState({ rows: newRows }, () => this.props.onChange(JSON.stringify(this.state.rows)));
    } catch (error:any) {
      void LoggerService.log('RepeaterInput-addRow', 'Medium', 'UserAction', error instanceof Error ? error.message : String(error));
    }
  }

  private _removeRow = (index: number): void => {
    try {
      const newRows = [...this.state.rows];
      newRows.splice(index, 1);
      this.setState({ rows: newRows }, () => this.props.onChange(JSON.stringify(newRows)));
    } catch (error:any) {
      void LoggerService.log('RepeaterInput-removeRow', 'Medium', 'UserAction', error instanceof Error ? error.message : String(error));
    }
  }

  private _validateDate(val: string, rule: string, n: number): string | null {
    try {
      if (!val) return null;
      const inputDate = new Date(val);
      inputDate.setHours(0, 0, 0, 0);
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      if (isNaN(inputDate.getTime())) return "Invalid Date";

      if (rule === 'future' && inputDate <= today) return "Must be after today";
      if (rule === 'past' && inputDate >= today) return "Must be before today";
      
      if (rule === 'future_n') {
        const target = new Date(today.getTime());
        target.setDate(today.getDate() + (n || 0));
        if (inputDate <= target) return `Must be after ${target.toLocaleDateString()}`;
      }
      
      if (rule === 'past_n') {
        const target = new Date(today.getTime());
        target.setDate(today.getDate() - (n || 0));
        if (inputDate >= target) return `Must be before ${target.toLocaleDateString()}`;
      }
      return null;
    } catch {
      return null;
    }
  }

  public render(): React.ReactElement<IRepeaterInputProps> {
    try {
      const { columns, mode } = this.props;
      const { rows } = this.state;

      const getOptions = (optString: string): string[] => {
        if (!optString) return [];
        return optString.split(',').map(s => s.trim());
      };

      return (
        <div style={{ overflowX: 'auto', border: '1px solid #e1dfdd', borderRadius: '2px', background: '#fff' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', minWidth: '400px' }}>
            <thead>
              <tr style={{ background: '#f3f2f1', borderBottom: '1px solid #edebe9' }}>
                {columns.map((c, i) => (
                  <th key={i} style={{ padding: '8px 10px', textAlign: 'left', fontWeight: 600, color: '#323130' }}>
                    {c.name} {c.required && <span style={{ color: '#a4262c' }}>*</span>}
                  </th>
                ))}
                {mode === 'edit' && <th style={{ width: 40 }}></th>}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, rIdx) => (
                <tr key={rIdx} style={{ borderBottom: '1px solid #f3f2f1' }}>
                  {columns.map((c, cIdx) => {
                    const val = row[c.key];
                    let isInvalid = c.required && (val === undefined || val === null || val === '' || (Array.isArray(val) && val.length === 0));
                    let errorMsg = isInvalid ? 'Required' : '';

                    if (!isInvalid && c.type === 'date' && c.dateRule && val) {
                      const dateErr = this._validateDate(val, c.dateRule, c.dateDays || 0);
                      if (dateErr) { isInvalid = true; errorMsg = dateErr; }
                    }

                    if (!isInvalid && c.unique && val !== undefined && val !== null && val !== '') {
                      const duplicate = rows.some((r, i) => i !== rIdx && r[c.key] === val);
                      if (duplicate) { isInvalid = true; errorMsg = 'Must be unique'; }
                    }

                    const inputStyle = isInvalid ? { ...styles.input, ...styles.errorBorder } : styles.input;
                    const selectStyle = isInvalid ? { ...styles.select, ...styles.errorBorder } : styles.select;

                    return (
                      <td key={cIdx} style={{ padding: '8px', verticalAlign: 'top' }}>
                        {mode === 'edit' ? (
                          <div>
                            {(!c.type || c.type === 'text') && (
                              <input type="text" value={val || ''} onChange={(e: React.ChangeEvent<HTMLInputElement>) => this._updateRow(rIdx, c.key, e.target.value)} style={inputStyle} />
                            )}
                            
                            {c.type === 'number' && (
                              <input type="number" value={val || ''} onChange={(e: React.ChangeEvent<HTMLInputElement>) => this._updateRow(rIdx, c.key, e.target.value)} style={inputStyle} />
                            )}

                            {c.type === 'date' && (
                              <input type="date" value={val || ''} onChange={(e: React.ChangeEvent<HTMLInputElement>) => this._updateRow(rIdx, c.key, e.target.value)} style={inputStyle} />
                            )}
                            
                            {c.type === 'choice' && (
                              <select value={val || ''} onChange={(e: React.ChangeEvent<HTMLSelectElement>) => this._updateRow(rIdx, c.key, e.target.value)} style={selectStyle}>
                                <option value="">Select...</option>
                                {getOptions(c.options || '').map((opt, k) => <option key={k} value={opt}>{opt}</option>)}
                              </select>
                            )}
                            
                            {c.type === 'multichoice' && (
                              <select 
                                multiple 
                                value={Array.isArray(val) ? val : []} 
                                style={{ ...selectStyle, height: '80px' }} 
                                onChange={(e: React.ChangeEvent<HTMLSelectElement>) => {
                                  const options = e.target.options;
                                  const selected: string[] = [];
                                  for (let i = 0; i < options.length; i++) {
                                    if (options[i].selected) selected.push(options[i].value);
                                  }
                                  this._updateRow(rIdx, c.key, selected);
                                }}
                              >
                                {getOptions(c.options || '').map((opt, k) => <option key={k} value={opt}>{opt}</option>)}
                              </select>
                            )}
                            {isInvalid && <span style={styles.errorText}>{errorMsg}</span>}
                          </div>
                        ) : (
                          <span>{Array.isArray(row[c.key]) ? row[c.key].join(', ') : (row[c.key] || '-')}</span>
                        )}
                      </td>
                    );
                  })}
                  {mode === 'edit' && (
                    <td style={{ textAlign: 'center', verticalAlign: 'middle', padding: '8px' }}>
                      <IconButton iconProps={{ iconName: 'Delete' }} styles={{ root: { color: '#a4262c' } }} onClick={() => this._removeRow(rIdx)} />
                    </td>
                  )}
                </tr>
              ))}
              {rows.length === 0 && <tr><td colSpan={columns.length + 1} style={{ padding: 20, textAlign: 'center', color: '#605e5c' }}>No items added.</td></tr>}
            </tbody>
          </table>
          {mode === 'edit' && (
            <div style={{ padding: '8px', background: '#faf9f8', borderTop: '1px solid #edebe9' }}>
              <button type="button" onClick={this._addRow} style={{ border: 'none', background: 'transparent', color: '#0078d4', cursor: 'pointer', fontWeight: 600, fontSize: '13px', display: 'flex', alignItems: 'center' }}>
                <Icon iconName="Add" style={{ marginRight: 6 }} /> Add Row
              </button>
            </div>
          )}
        </div>
      );
    } catch (error:any) {
      void LoggerService.log('RepeaterInput-render', 'High', 'UI', error instanceof Error ? error.message : String(error));
      return <div style={{ color: 'red', padding: 10 }}>Error rendering grid.</div>;
    }
  }
}