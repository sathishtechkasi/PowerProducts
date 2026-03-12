import * as React from 'react';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LoggerService } from './LoggerService';

export interface IViewFilter {
  field: string;
  operator: 'eq' | 'ne' | 'gt' | 'lt' | 'geq' | 'leq' | 'contains';
  value: string;
}

export interface IViewConfig {
  id: string;
  title: string;
  allowedGroups: string[];
  filters: IViewFilter[];
  visibleFields: string[];
}

export interface IViewEditorProps {
  views: IViewConfig[];
  fields: { key: string; text: string }[];
  context: WebPartContext;
  onSave: (views: IViewConfig[]) => void;
  onCancel: () => void;
}

export interface IViewEditorState {
  views: IViewConfig[];
  editing: Partial<IViewConfig>;
  siteGroups: { Id: number, Title: string }[];
  loadingGroups: boolean;
}

export class ViewEditor extends React.Component<IViewEditorProps, IViewEditorState> {
  constructor(props: IViewEditorProps) {
    super(props);
    try {
      this.state = {
        views: props.views ? JSON.parse(JSON.stringify(props.views)) : [],
        editing: this._getEmptyView(),
        siteGroups: [],
        loadingGroups: true
      };
    } catch (error:any) {
      void LoggerService.log('ViewEditor-constructor', 'High', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  public componentDidMount(): void {
    void this._loadGroups();
  }

  private _getEmptyView(): Partial<IViewConfig> {
    return { id: '', title: '', allowedGroups: [], filters: [], visibleFields: [] };
  }

  private async _loadGroups(): Promise<void> {
    try {
      const response = await this.props.context.spHttpClient.get(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/sitegroups?$select=Id,Title&$orderby=Title`,
        SPHttpClient.configurations.v1
      );
      const data = await response.json();
      this.setState({ siteGroups: data.value, loadingGroups: false });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-loadGroups', 'High', 'Config', error instanceof Error ? error.message : String(error));
      this.setState({ loadingGroups: false });
    }
  }

  private _saveView = (): void => {
    try {
      const { editing, views } = this.state;
      if (!editing.title) return;
      
      const finalFields = (editing.visibleFields && editing.visibleFields.length > 0) 
        ? editing.visibleFields 
        : ['Title', 'Created', 'Modified'];
        
      const newView = {
        ...editing,
        id: editing.id || `view_${Date.now()}`,
        visibleFields: finalFields
      } as IViewConfig;
      
      const isUpdate = views.some(v => v.id === newView.id);
      const updatedViews = isUpdate ? views.map(v => v.id === newView.id ? newView : v) : [...views, newView];
      
      this.setState({ views: updatedViews, editing: this._getEmptyView() });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-saveView', 'High', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _editView = (view: IViewConfig): void => {
    try {
      this.setState({ editing: JSON.parse(JSON.stringify(view)) });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-editView', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _deleteView = (id: string): void => {
    try {
      if(!confirm('Delete this view?')) return;
      this.setState(prev => ({
        views: prev.views.filter(v => v.id !== id),
        editing: prev.editing.id === id ? this._getEmptyView() : prev.editing
      }));
    } catch (error:any) {
      void LoggerService.log('ViewEditor-deleteView', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _addColumn = (fieldKey?: string): void => {
    try {
      if (!fieldKey) return;
      const current = this.state.editing.visibleFields || [];
      if (current.indexOf(fieldKey) === -1) {
        this.setState({ editing: { ...this.state.editing, visibleFields: [...current, fieldKey] } });
      }
    } catch (error:any) {
      void LoggerService.log('ViewEditor-addColumn', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _removeColumn = (index: number): void => {
    try {
      const current = [...(this.state.editing.visibleFields || [])];
      current.splice(index, 1);
      this.setState({ editing: { ...this.state.editing, visibleFields: current } });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-removeColumn', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _moveColumn = (index: number, direction: 'up' | 'down'): void => {
    try {
      const current = [...(this.state.editing.visibleFields || [])];
      if (direction === 'up' && index > 0) {
        [current[index], current[index - 1]] = [current[index - 1], current[index]];
      } else if (direction === 'down' && index < current.length - 1) {
        [current[index], current[index + 1]] = [current[index + 1], current[index]];
      }
      this.setState({ editing: { ...this.state.editing, visibleFields: current } });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-moveColumn', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _addFilter = (): void => {
    try {
      const currentFilters = this.state.editing.filters || [];
      this.setState({
        editing: { ...this.state.editing, filters: [...currentFilters, { field: '', operator: 'eq', value: '' }] }
      });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-addFilter', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _updateFilter = (index: number, key: keyof IViewFilter, val: string): void => {
    try {
      const newFilters = [...(this.state.editing.filters || [])];
      newFilters[index] = { ...newFilters[index], [key]: val };
      this.setState({ editing: { ...this.state.editing, filters: newFilters } });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-updateFilter', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _removeFilter = (index: number): void => {
    try {
      const newFilters = [...(this.state.editing.filters || [])];
      newFilters.splice(index, 1);
      this.setState({ editing: { ...this.state.editing, filters: newFilters } });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-removeFilter', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _toggleGroup = (groupName: string, checked: boolean): void => {
    try {
      const current = this.state.editing.allowedGroups || [];
      let newGroups = [...current];
      if (checked) {
        if (newGroups.indexOf(groupName) === -1) newGroups.push(groupName);
      } else {
        newGroups = newGroups.filter(g => g !== groupName);
      }
      this.setState({ editing: { ...this.state.editing, allowedGroups: newGroups } });
    } catch (error:any) {
      void LoggerService.log('ViewEditor-toggleGroup', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  public render(): React.ReactElement<IViewEditorProps> {
    const { views, editing, siteGroups } = this.state;
    const isEditing = !!editing.id;

    const getFieldName = (key: string) => {
      const f = this.props.fields.find(x => x.key === key);
      return f ? f.text : key;
    };

    const fieldOptions: IDropdownOption[] = this.props.fields.map(f => ({ key: f.key, text: f.text }));

    return (
      <div style={{ fontSize: 13, padding: '10px 0' }}>
        <h3 style={{ marginTop: 0, color: '#323130' }}>Custom Views ({views.length})</h3>
        
        {/* --- LIST OF EXISTING VIEWS --- */}
        <div style={{ maxHeight: 150, overflowY: 'auto', border: '1px solid #e1dfdd', marginBottom: 15, background: '#fff', borderRadius: 2 }}>
          {views.map(v => (
            <div key={v.id} style={{ padding: '8px 12px', borderBottom: '1px solid #f3f2f1', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: editing.id === v.id ? '#eff6ff' : 'transparent' }}>
              <span style={{ fontWeight: 600 }}>{v.title}</span>
              <div style={{ display: 'flex', gap: 5 }}>
                <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editView(v)} styles={{ root: { color: '#0078d4', height: 24, width: 24 } }} />
                <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._deleteView(v.id)} styles={{ root: { color: '#d13438', height: 24, width: 24 } }} />
              </div>
            </div>
          ))}
          {views.length === 0 && <div style={{ padding: 15, color: '#605e5c', fontStyle: 'italic', textAlign: 'center' }}>No views defined.</div>}
        </div>

        {/* --- ADD / EDIT FORM --- */}
        <div style={{ border: '1px solid #e1dfdd', padding: 15, background: isEditing ? '#fdfdfd' : '#faf9f8', borderRadius: 2 }}>
          <h4 style={{ margin: '0 0 15px', color: isEditing ? '#0078d4' : '#323130' }}>{isEditing ? 'Edit View' : 'Create View'}</h4>
          
          <TextField 
            label="View Title" 
            required 
            value={editing.title || ''} 
            onChange={(_, val) => this.setState({ editing: { ...editing, title: val || '' } })}
            placeholder="e.g. My Open Items"
            styles={{ root: { marginBottom: 15 } }}
          />

          <div style={{ marginBottom: 15 }}>
            <label style={{ display: 'block', fontWeight: 600, marginBottom: 5 }}>Columns & Order</label>
            <div style={{ background: '#fff', border: '1px solid #edebe9', padding: 10, borderRadius: 2 }}>
              <Dropdown
                placeholder="+ Add Column..."
                options={fieldOptions}
                onChange={(_, opt) => this._addColumn(opt?.key as string)}
                styles={{ root: { marginBottom: 10 } }}
              />
              
              <div style={{ maxHeight: 150, overflowY: 'auto' }}>
                 {(editing.visibleFields || []).map((colKey, idx) => (
                    <div key={colKey} style={{ display: 'flex', alignItems: 'center', padding: '6px', borderBottom: '1px solid #f3f2f1' }}>
                       <span style={{ flex: 1 }}>{getFieldName(colKey)}</span>
                       <IconButton iconProps={{ iconName: 'ChevronUpSmall' }} disabled={idx === 0} onClick={() => this._moveColumn(idx, 'up')} styles={{ root: { height: 24, width: 24 } }} />
                       <IconButton iconProps={{ iconName: 'ChevronDownSmall' }} disabled={idx === (editing.visibleFields || []).length - 1} onClick={() => this._moveColumn(idx, 'down')} styles={{ root: { height: 24, width: 24 } }} />
                       <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => this._removeColumn(idx)} styles={{ root: { color: '#d13438', height: 24, width: 24 } }} />
                    </div>
                 ))}
                 {(!editing.visibleFields || editing.visibleFields.length === 0) && (
                   <div style={{ color: '#a19f9d', fontStyle: 'italic', padding: 5, fontSize: 12 }}>No columns selected. Default columns will be used.</div>
                 )}
              </div>
            </div>
          </div>

          <div style={{ marginBottom: 15 }}>
            <label style={{ display: 'block', fontWeight: 600, marginBottom: 5 }}>Permissions (Optional)</label>
            <div style={{ fontSize: 11, color: '#605e5c', marginBottom: 5 }}>Visible to:</div>
            <div style={{ maxHeight: 100, overflowY: 'auto', border: '1px solid #edebe9', background: 'white', padding: '8px 10px' }}>
              {siteGroups.map(g => (
                <Checkbox 
                  key={g.Id} 
                  label={g.Title} 
                  checked={(editing.allowedGroups || []).indexOf(g.Title) > -1}
                  onChange={(_, checked) => this._toggleGroup(g.Title, !!checked)}
                  styles={{ root: { marginBottom: 6 } }}
                />
              ))}
            </div>
          </div>

          <div style={{ marginBottom: 15 }}>
            <label style={{ display: 'block', fontWeight: 600, marginBottom: 5 }}>Filter Criteria</label>
            <div style={{ background: '#fff', border: '1px solid #edebe9', padding: 10, borderRadius: 2 }}>
              {(editing.filters || []).map((f, i) => (
                <div key={i} style={{ display: 'flex', gap: 10, marginBottom: 10, alignItems: 'flex-end' }}>
                  <div style={{ flex: 2 }}>
                    <Dropdown 
                      options={fieldOptions} 
                      selectedKey={f.field} 
                      onChange={(_, opt) => this._updateFilter(i, 'field', opt?.key as string)} 
                    />
                  </div>
                  <div style={{ flex: 1 }}>
                    <Dropdown 
                      options={[
                        { key: 'eq', text: '=' }, { key: 'contains', text: 'Has' }, 
                        { key: 'geq', text: '>=' }, { key: 'leq', text: '<=' }
                      ]} 
                      selectedKey={f.operator} 
                      onChange={(_, opt) => this._updateFilter(i, 'operator', opt?.key as string)} 
                    />
                  </div>
                  <div style={{ flex: 2 }}>
                    <TextField value={f.value} onChange={(_, val) => this._updateFilter(i, 'value', val || '')} placeholder="Value" />
                  </div>
                  <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._removeFilter(i)} styles={{ root: { color: '#d13438' } }} />
                </div>
              ))}
              <DefaultButton text="+ Add Filter" onClick={this._addFilter} styles={{ root: { color: '#0078d4' } }} />
            </div>
          </div>

          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
             <DefaultButton text="Cancel" onClick={() => this.setState({ editing: this._getEmptyView() })} />
             <PrimaryButton 
                text={isEditing ? 'Update View' : 'Add View'} 
                disabled={!editing.title}
                onClick={this._saveView} 
             />
          </div>
        </div>

        {/* --- MAIN FOOTER BUTTONS --- */}
        <div style={{ marginTop: 20, paddingTop: 15, borderTop: '1px solid #edebe9', textAlign: 'right', display: 'flex', justifyContent: 'flex-end', gap: 10 }}>
            <PrimaryButton text="Save All Changes" onClick={() => this.props.onSave(this.state.views)} />
            <DefaultButton text="Close" onClick={this.props.onCancel} />
        </div>
      </div>
    );
  }
}