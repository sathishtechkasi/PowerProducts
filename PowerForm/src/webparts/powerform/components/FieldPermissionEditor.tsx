import * as React from 'react';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LoggerService } from './LoggerService';

export interface IFieldPermissionEditorProps {
  fieldKey: string;
  fieldTitle: string;
  context: WebPartContext; 
  selectedGroups: string[]; 
  onSave: (groups: string[]) => void;
  onCancel: () => void;
}

export interface IFieldPermissionEditorState {
  loading: boolean;
  siteGroups: { Id: number, Title: string }[];
  selected: string[];
}

export class FieldPermissionEditor extends React.Component<IFieldPermissionEditorProps, IFieldPermissionEditorState> {
  constructor(props: IFieldPermissionEditorProps) {
    super(props);
    try {
      this.state = {
        loading: true,
        siteGroups: [],
        selected: props.selectedGroups || []
      };
    } catch (error:any) {
      void LoggerService.log('FieldPermissionEditor-constructor', 'High', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  public componentDidMount(): void {
    void this._loadGroups();
  }

  private async _loadGroups(): Promise<void> {
    try {
      const response = await this.props.context.spHttpClient.get(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/sitegroups?$select=Id,Title&$orderby=Title`,
        SPHttpClient.configurations.v1
      );
      
      if (!response.ok) {
        throw new Error(`SharePoint API failed: ${response.statusText}`);
      }
      
      const data = await response.json();
      this.setState({ siteGroups: data.value, loading: false });
    } catch (error:any) {
      void LoggerService.log('FieldPermissionEditor-loadGroups', 'High', 'Config', error instanceof Error ? error.message : String(error));
      this.setState({ loading: false });
    }
  }

  private _onToggleGroup = (groupName: string, checked: boolean): void => {
    try {
      this.setState(prev => {
        const newSelected = [...prev.selected];
        if (checked) {
          if (newSelected.indexOf(groupName) === -1) newSelected.push(groupName);
        } else {
          const idx = newSelected.indexOf(groupName);
          if (idx > -1) newSelected.splice(idx, 1);
        }
        return { selected: newSelected };
      });
    } catch (error:any) {
      void LoggerService.log('FieldPermissionEditor-onToggleGroup', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  public render(): React.ReactElement<IFieldPermissionEditorProps> {
    if (this.state.loading) {
      return (
        <div style={{ padding: 20, textAlign: 'center' }}>
          <Spinner size={SpinnerSize.small} label="Loading SharePoint groups..." />
        </div>
      );
    }

    return (
      <div style={{ padding: 15, border: '1px solid #e1dfdd', background: '#faf9f8', marginBottom: 10, borderRadius: 2 }}>
        <h4 style={{ marginTop: 0, marginBottom: 10, color: '#323130', fontSize: '14px' }}>
          Permissions for "{this.props.fieldTitle}"
        </h4>
        
        <p style={{ fontSize: 12, color: '#605e5c', marginBottom: 15, lineHeight: '1.4' }}>
          Select groups allowed to <b>EDIT</b> this field. <br/>
          If no groups are selected, <b>everyone</b> can edit it.
          <br/>Users not in these groups will see it as <b>Read Only</b>.
        </p>

        <div style={{ maxHeight: 200, overflowY: 'auto', border: '1px solid #edebe9', background: '#ffffff', padding: '10px 5px' }}>
          {this.state.siteGroups.length === 0 ? (
            <div style={{ fontSize: 12, color: '#a19f9d', textAlign: 'center', padding: 10 }}>No groups found.</div>
          ) : (
            this.state.siteGroups.map(g => (
              <Checkbox
                key={g.Id}
                label={g.Title}
                checked={this.state.selected.indexOf(g.Title) > -1}
                onChange={(_, checked) => this._onToggleGroup(g.Title, !!checked)}
                styles={{ root: { marginBottom: 8, paddingLeft: 5 } }}
              />
            ))
          )}
        </div>

        <div style={{ marginTop: 15, display: 'flex', gap: 10 }}>
          <PrimaryButton 
            text="Save"
            onClick={() => {
              try {
                this.props.onSave(this.state.selected);
              } catch (error:any) {
                void LoggerService.log('FieldPermissionEditor-onSave', 'High', 'Config', error instanceof Error ? error.message : String(error));
              }
            }}
          />
          <DefaultButton 
            text="Cancel"
            onClick={this.props.onCancel}
          />
        </div>
      </div>
    );
  }
}