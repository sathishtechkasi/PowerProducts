import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { DefaultButton, PrimaryButton, IconButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { IFormSection } from './IFormSection';
import { LoggerService } from './LoggerService';
import Swal from 'sweetalert2';

export interface ISectionEditorProps {
  sections: IFormSection[];
  availableFields: { key: string; text: string }[];
  onSave: (sections: IFormSection[]) => void;
  onCancel: () => void;
}

export interface ISectionEditorState {
  sections: IFormSection[];
}

export class SectionEditor extends React.Component<ISectionEditorProps, ISectionEditorState> {
  constructor(props: ISectionEditorProps) {
    super(props);
    try {
      this.state = {
        sections: props.sections ? JSON.parse(JSON.stringify(props.sections)) : []
      };
    } catch (error:any) {
      void LoggerService.log('SectionEditor-constructor', 'High', 'Config', error instanceof Error ? error.message : String(error));
      this.state = { sections: [] };
    }
  }

  private _addSection = (): void => {
    try {
      const newSection: IFormSection = {
        id: 'section_' + new Date().getTime(),
        title: '',
        fields: [],
        order: this.state.sections.length + 1,
        columns: 1,
        isCollapsible: false
      };
      this.setState({ sections: [...this.state.sections, newSection] });
    } catch (error:any) {
      void LoggerService.log('SectionEditor-addSection', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _removeSection = (index: number): void => {
    try {
      const newSec = [...this.state.sections];
      newSec.splice(index, 1);
      this.setState({ sections: newSec });
    } catch (error:any) {
      void LoggerService.log('SectionEditor-removeSection', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _updateSection = (index: number, key: keyof IFormSection, val: any): void => {
    try {
      const newSec = [...this.state.sections];
      newSec[index] = { ...newSec[index], [key]: val };
      this.setState({ sections: newSec });
    } catch (error:any) {
      void LoggerService.log('SectionEditor-updateSection', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  private _moveSection = (index: number, direction: 'up' | 'down'): void => {
    try {
      const newSec = [...this.state.sections];
      if (direction === 'up' && index > 0) {
        [newSec[index], newSec[index - 1]] = [newSec[index - 1], newSec[index]];
      } else if (direction === 'down' && index < newSec.length - 1) {
        [newSec[index], newSec[index + 1]] = [newSec[index + 1], newSec[index]];
      }
      this.setState({ sections: newSec });
    } catch (error:any) {
      void LoggerService.log('SectionEditor-moveSection', 'Low', 'Config', error instanceof Error ? error.message : String(error));
    }
  }

  public render(): React.ReactElement<ISectionEditorProps> {
    const allUsedFields = new Set<string>();
    try {
      this.state.sections.forEach(s => {
        if (s.fields) {
          s.fields.forEach(f => allUsedFields.add(f));
        }
      });
    } catch (error:any) {
      void LoggerService.log('SectionEditor-render-calc', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
    }

    return (
      <Panel
        isOpen={true}
        onDismiss={this.props.onCancel}
        type={PanelType.medium}
        headerText="Form Sections Configuration"
        closeButtonAriaLabel="Close"
      >
        <div style={{ marginBottom: 20 }}>
          <p style={{ color: '#605e5c', fontSize: '13px' }}>
            Create sections to group fields. Fields can only belong to one section at a time.
          </p>
          <PrimaryButton text="+ Add Section" onClick={this._addSection} />
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: '15px', paddingBottom: 50 }}>
          {this.state.sections.map((section, idx) => {
            const availableOptions = this.props.availableFields.filter(f => {
              const isUsedElsewhere = allUsedFields.has(f.key) && section.fields.indexOf(f.key) === -1;
              return !isUsedElsewhere;
            });

            return (
              <div key={section.id} style={{ border: '1px solid #e1dfdd', padding: 15, borderRadius: 4, background: '#faf9f8' }}>
                
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 10 }}>
                  <div style={{ flex: 1, marginRight: 15 }}>
                    <TextField
                      label="Section Title"
                      value={section.title}
                      placeholder="e.g., General Information"
                      onChange={(_, val) => this._updateSection(idx, 'title', val || '')}
                      required
                    />
                    <TextField
                      label="Description (Optional)"
                      value={section.description || ''}
                      placeholder="Help text shown below the title"
                      onChange={(_, val) => this._updateSection(idx, 'description', val || '')}
                    />
                  </div>
                  
                  <div style={{ display: 'flex', marginTop: 28 }}>
                    <IconButton iconProps={{ iconName: 'ChevronUp' }} title="Move Up" disabled={idx === 0} onClick={() => this._moveSection(idx, 'up')} />
                    <IconButton iconProps={{ iconName: 'ChevronDown' }} title="Move Down" disabled={idx === this.state.sections.length - 1} onClick={() => this._moveSection(idx, 'down')} />
                    <IconButton iconProps={{ iconName: 'Delete' }} title="Delete Section" onClick={() => this._removeSection(idx)} styles={{ root: { color: '#d13438' }, rootHovered: { background: '#fde7e9' } }} />
                  </div>
                </div>

                <div style={{ display: 'flex', gap: '20px', marginBottom: 15, alignItems: 'flex-end' }}>
                  <div style={{ width: '150px' }}>
                    <Dropdown
                      label="Column Layout"
                      options={[
                        { key: 1, text: '1 Column (Full Width)' },
                        { key: 2, text: '2 Columns (50/50)' },
                        { key: 3, text: '3 Columns (33/33/33)' }
                      ]}
                      selectedKey={section.columns || 1}
                      onChange={(_, opt) => this._updateSection(idx, 'columns', opt?.key)}
                    />
                  </div>
                  <Toggle
                    label="Collapsible Section?"
                    inlineLabel
                    checked={!!section.isCollapsible}
                    onChange={(_, checked) => this._updateSection(idx, 'isCollapsible', checked)}
                    styles={{ root: { marginBottom: 0 } }}
                  />
                </div>

                <Dropdown
                  label="Select Fields for this Section"
                  multiSelect
                  options={availableOptions.map(f => ({ key: f.key, text: f.text }))}
                  selectedKeys={section.fields}
                  onChange={(_, opt?: IDropdownOption) => {
                    try {
                      if (opt) {
                        const current = section.fields || [];
                        let next: string[] = [];
                        if (opt.selected) next = [...current, opt.key as string];
                        else next = current.filter(k => k !== opt.key);
                        this._updateSection(idx, 'fields', next);
                      }
                    } catch (error:any) {
                      void LoggerService.log('SectionEditor-dropdown', 'Medium', 'Config', error instanceof Error ? error.message : String(error));
                    }
                  }}
                />
              </div>
            );
          })}
        </div>

        <div style={{ marginTop: 20, paddingTop: 20, borderTop: '1px solid #edebe9' }}>
          <PrimaryButton
            text="Save Configuration"
            onClick={() => {
              try {
                const sections = this.state.sections;
                const titles = sections.map(s => (s.title || '').trim());
                
                if (titles.some(t => t === '')) {
                  void Swal.fire({ icon: 'warning', title: 'Warning', text: 'Please provide a title for all sections.' });
                  return;
                }
                
                const uniqueTitles = new Set(titles);
                if (uniqueTitles.size !== titles.length) {
                  void Swal.fire({ icon: 'warning', title: 'Warning', text: 'Section titles must be unique. Please rename duplicates.' });
                  return;
                }
                
                this.props.onSave(this.state.sections);
              } catch (error:any) {
                void LoggerService.log('SectionEditor-onSave', 'High', 'Config', error instanceof Error ? error.message : String(error));
              }
            }}
            style={{ marginRight: 10 }}
          />
          <DefaultButton text="Cancel" onClick={this.props.onCancel} />
        </div>
      </Panel>
    );
  }
}