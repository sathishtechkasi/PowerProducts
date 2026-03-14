import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

// UPGRADE: PnP v4 Imports
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

// UPGRADE: Modern SweetAlert import
import Swal from 'sweetalert2';

export interface IAdminPanelProps {
  isOpen: boolean;
  onClose: () => void;
  listName: string;
  sp: SPFI; // UPGRADE: Accept the configured SPFI instance from the Customizer
}

export interface IRule {
  Id?: number;
  Title: string;
  Selector: string;
  ActionType: string;
  CSSValue: string;
  IsOverride: boolean;
  IsActive: boolean;
}

export interface IAdminPanelState {
  rules: IRule[];
  loading: boolean;
  importing: boolean;
  isEditing: boolean;
  currentRule: IRule;
  expandedGroups: { active: boolean; inactive: boolean };
  expandedSubGroups: { [key: string]: boolean };
  isExporting?: boolean;
  exportJson?: string;
  exportJsonName?: string;
  isHelpOpen?: boolean;
}

export default class AdminPanel extends React.Component<IAdminPanelProps, IAdminPanelState> {
  private fileInput: HTMLInputElement | null;
  private Toast: any;

  constructor(props: IAdminPanelProps) {
    super(props);
    this.fileInput = null;
    this.triggerImport = this.triggerImport.bind(this);
    this.handleFileChange = this.handleFileChange.bind(this);
    this.handleExport = this.handleExport.bind(this);
    this.saveRule = this.saveRule.bind(this);
    this.deleteRule = this.deleteRule.bind(this);
    this.renderForm = this.renderForm.bind(this);
    this.renderList = this.renderList.bind(this);
    this.renderGroupedRules = this.renderGroupedRules.bind(this);
    this.setFileInputRef = this.setFileInputRef.bind(this);
    
    this.state = {
      rules: [],
      loading: true,
      importing: false,
      isEditing: false,
      currentRule: this.getEmptyRule(),
      expandedGroups: { active: true, inactive: true },
      expandedSubGroups: {},
      isExporting: false,
      exportJson: '',
      exportJsonName: '',
      isHelpOpen: false
    };

    // UPGRADE: 'onOpen' is deprecated in SweetAlert2 v11+, using 'didOpen' instead
    this.Toast = Swal.mixin({
      toast: true,
      position: 'top-end',
      showConfirmButton: false,
      timer: 4000,
      timerProgressBar: false,
      didOpen: (toast: HTMLElement) => {
        const container = toast.parentElement;
        if (container) {
          container.style.zIndex = '2147483647';
          container.style.position = 'fixed';
        }
      }
    });
  }

  public componentDidMount() {
    void this.fetchRules();
  }

  private setFileInputRef(element: HTMLInputElement) {
    this.fileInput = element;
  }

  private getEmptyRule(): IRule {
    return { Title: '', Selector: '', ActionType: 'Hide', CSSValue: '', IsOverride: false, IsActive: true };
  }

  private async fetchRules() {
    this.setState({ loading: true });
    try {
      // UPGRADE: Use this.props.sp and execute with () instead of .get()
      const items = await this.props.sp.web.lists.getByTitle(this.props.listName).items.top(5000)();
      this.setState({ rules: items, loading: false });
    } catch (error) {
      console.error(error);
      this.setState({ loading: false });
      void this.Toast.fire({
        icon: 'error',
        title: 'Fetch Failed',
        text: 'Could not fetch rules...'
      });
    }
  }

  // --- EXPORT ---
  // --- UI-BASED EXPORT (TEXTAREA IN PANEL - ALWAYS COPYABLE) ---
  private handleExport() {
    const rules = this.state.rules;
    if (!rules || rules.length === 0) {
      alert('No Data: There are no rules to export.');
      return;
    }

    // Generate both CSV and JSON
    const escapeCSV = (value: any): string => {
      let str = (value === null || value === undefined) ? '' : value.toString();
      if (str.indexOf(',') > -1 || str.indexOf('"') > -1 || str.indexOf('\n') > -1 || str.indexOf('\r') > -1) {
        str = '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    };

    const csvLines: string[] = [];
    csvLines.push(['Title', 'Selector', 'ActionType', 'CSSValue', 'IsOverride', 'IsActive'].map(escapeCSV).join(','));
    
    rules.forEach((rule: any) => {
      csvLines.push([
        rule.Title || '',
        rule.Selector || '',
        rule.ActionType || 'Hide',
        rule.CSSValue || '',
        rule.IsOverride || false,
        rule.IsActive || false
      ].map(escapeCSV).join(','));
    });

    const cleanRules = rules.map((r: any) => {
      return {
        Title: r.Title || "",
        Selector: r.Selector || "",
        ActionType: r.ActionType || "Hide",
        CSSValue: r.CSSValue || "",
        IsOverride: r.IsOverride || false,
        IsActive: r.IsActive || false
      };
    });

    const json = JSON.stringify(cleanRules, null, 2);
    const fileNameJson = "BAZ_Config_" + new Date().getTime() + ".json";

    this.setState({
      isEditing: false,
      exportJson: json,
      exportJsonName: fileNameJson,
      isExporting: true
    });
  }

  // --- IMPORT ---
  private triggerImport() {
    if (this.fileInput) {
      this.fileInput.click();
    }
  }

  private handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const target = e.target as any;
        const text = target.result as string;
        const json = JSON.parse(text);
        if (Array.isArray(json)) {
          void this.processImport(json);
        } else {
          throw new Error("Invalid JSON");
        }
      } catch (err) {
        void Swal.fire({
          toast: true,
          position: 'top-end',
          icon: 'error',
          title: 'Import Failed',
          text: 'Invalid JSON file.',
          showConfirmButton: false,
          timer: 4000
        });
      }
      event.target.value = '';
    };
    reader.readAsText(file);
  }

  private async processImport(importedRules: IRule[]) {
    void this.Toast.fire({
      icon: 'warning',
      title: 'Import Started',
      text: 'Scanning and importing items... Please wait'
    });
    this.setState({ importing: true });
    
    const list = this.props.sp.web.lists.getByTitle(this.props.listName);
    let successCount = 0;
    let updateCount = 0;
    let failCount = 0;

    for (let i = 0; i < importedRules.length; i++) {
      const rule = importedRules[i];
      try {
        const escapedTitle = rule.Title.replace(/'/g, "''");
        // UPGRADE: Use () for execution
        const existingItems = await list.items.filter("Title eq '" + escapedTitle + "'")();
        const payload = {
          Title: rule.Title,
          Selector: rule.Selector,
          ActionType: rule.ActionType,
          CSSValue: rule.CSSValue,
          IsOverride: rule.IsOverride,
          IsActive: rule.IsActive
        };
        if (existingItems && existingItems.length > 0) {
          const idToUpdate = existingItems[0].Id;
          await list.items.getById(idToUpdate).update(payload);
          updateCount++;
        } else {
          await list.items.add(payload);
          successCount++;
        }
      } catch (e) {
        failCount++;
      }
    }
    
    this.setState({ importing: false });

    void this.Toast.fire({
      icon: 'success',
      title: 'Import Complete',
      text: `Created: ${successCount}, Updated: ${updateCount}`
    });
    void this.fetchRules();
  }

  // --- CRUD ---
  private async saveRule() {
    const currentRule = this.state.currentRule;
    const list = this.props.sp.web.lists.getByTitle(this.props.listName);

    if (!currentRule.Title) {
      void this.Toast.fire({ icon: 'error', title: 'Validation Error', text: 'Rule Name is required.' });
      return;
    }

    // UNIQUE LOADER CHECK
    if (currentRule.ActionType === 'Loader' && currentRule.IsActive) {
      const rules = this.state.rules;
      const conflictMap = rules.map((r) => {
        return r.ActionType === 'Loader' && r.IsActive && r.Id !== currentRule.Id;
      });

      // Find the index of the first 'true'
      const activeLoaderIndex = conflictMap.indexOf(true);

      if (activeLoaderIndex !== -1) {
        const activeLoader = rules[activeLoaderIndex];
        void this.Toast.fire({
          icon: 'error',
          title: 'Config Error',
          text: 'Only one Loader can be active. Please deactivate "' + activeLoader.Title + '" first.'
        });
        return;
      }
    }
    // ----------------------------------------------

    // DUPLICATE NAME CHECK
    const duplicate = this.state.rules.filter((r) => {
      return r.Title.toLowerCase() === currentRule.Title.toLowerCase() && r.Id !== currentRule.Id;
    });

    if (duplicate.length > 0) {
      void this.Toast.fire({ icon: 'error', title: 'Duplicate Error', text: 'A rule with this Name already exists.' });
      return;
    }

    try {
      if (currentRule.Id) {
        await list.items.getById(currentRule.Id).update(currentRule);
      } else {
        await list.items.add(currentRule);
      }
      void this.Toast.fire({ icon: 'success', title: 'Success', text: 'Changes saved.' });
      this.setState({ isEditing: false, currentRule: this.getEmptyRule() });
      void this.fetchRules();
    } catch (error) {
      console.error(error);
      void this.Toast.fire({ icon: 'error', title: 'Save Failed', text: 'Server error occurred.' });
    }
  }

  // --- INSTANT TOGGLE ---
  private async toggleRuleActive(rule: IRule, checked: boolean) {
    const list = this.props.sp.web.lists.getByTitle(this.props.listName);

    // Optimistic Update
    const updatedRules = this.state.rules.map((r) => {
      if (r.Id === rule.Id) {
        return { ...r, IsActive: checked };
      }
      return r;
    });

    this.setState({ rules: updatedRules });

    try {
      await list.items.getById(rule.Id!).update({ IsActive: checked });
      void this.Toast.fire({
        icon: 'success',
        title: 'Updated',
        text: `Rule "${rule.Title}" is now ${checked ? 'Active' : 'Inactive'}`,
        timer: 1500
      });

    } catch (error) {
      console.error(error);
      // Revert if failed
      const revertedRules = this.state.rules.map((r) => {
        if (r.Id === rule.Id) {
          return { ...r, IsActive: !checked };
        }
        return r;
      });
      this.setState({ rules: revertedRules });
      void this.Toast.fire({ icon: 'error', title: 'Update Failed', text: 'Could not update status.' });
    }
  }

  private deleteRule(id: number) {
    if (confirm("Are you sure you want to delete this rule?")) {
      this.props.sp.web.lists.getByTitle(this.props.listName).items.getById(id).delete()
        .then(() => {
          void this.Toast.fire({ icon: 'success', title: 'Deleted', text: 'Rule removed.' });
          void this.fetchRules();
        })
        .catch((err) => {
          void this.Toast.fire({ icon: 'error', title: 'Delete Failed', text: 'Could not remove item.' });
        });
    }
  }

  // --- RENDER ---
  public render(): React.ReactElement<IAdminPanelProps> {
    return (
      <Panel
        isOpen={this.props.isOpen}
        type={PanelType.custom}
        customWidth="60%"
        onDismiss={this.props.onClose}
        onRenderHeader={() => {
          return (
            <div style={{ fontSize: '24px', fontWeight: 'bold', padding: '16px 20px' }}>
              BAZ Customizer Manager
            </div>
          );
        }}
      >
        <style>{`
          .swal2-container {
            z-index: 1000000 !important;  
          }
          .swal2-toast {
            z-index: 1000000 !important; 
          }
        `}</style>
        <input
          type="file"
          accept=".json"
          ref={this.setFileInputRef}
          style={{ display: 'none' }}
          onChange={this.handleFileChange}
        />
        {(this.state.loading || this.state.importing) ? (
          <div style={{ marginTop: '50px' }}>
            <Spinner size={SpinnerSize.large} label={this.state.importing ? "Importing Rules..." : "Loading..."} />
          </div>
        ) : (
          <div style={{ padding: '10px' }}>
            {this.state.isEditing ? this.renderForm() : this.renderList()}
          </div>
        )}
      </Panel>
    );
  }

  private renderList() {
    const rules = this.state.rules;
    const expandedGroups = this.state.expandedGroups;
    const activeRules = rules.filter((r) => r.IsActive);
    const inactiveRules = rules.filter((r) => !r.IsActive);
    const sectionHeaderStyle: React.CSSProperties = {
      display: 'flex', alignItems: 'center', cursor: 'pointer',
      borderBottom: '2px solid', marginTop: '20px', paddingBottom: '5px'
    };

    if (this.state.isExporting) {
      return (
        <div style={{ padding: '20px' }}>
          <h2>Export Config</h2>
          <p>Copy the text below and paste into a file.</p>
          <h3 style={{ marginTop: '40px' }}>JSON (for re-import)</h3>
          <p>Save as: <strong>{this.state.exportJsonName}</strong></p>
          <textarea
            style={{ width: '100%', height: '300px' }}
            value={this.state.exportJson}
            readOnly
          />
          <div style={{ marginTop: '20px' }}>
            <DefaultButton text="Back to List" onClick={() => this.setState({ isExporting: false })} />
          </div>
        </div>
      );
    }

    if (this.state.isHelpOpen) {
      return (
        <div style={{ padding: '20px', maxHeight: '85vh', overflowY: 'auto' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
            <h2 style={{ margin: 0 }}>How to Configure Rules</h2>
            <DefaultButton text="Back to List" onClick={() => this.setState({ isHelpOpen: false })} />
          </div>

          <div style={{ marginBottom: '20px', border: '1px solid #edebe9', padding: '15px', borderRadius: '4px', backgroundColor: '#faf9f8' }}>
            <h3 style={{ marginTop: 0, color: '#0078d4' }}>1. Hide (Removing an element)</h3>
            <p>Removes an element from the page. You only need the CSS selector.</p>
            <ul style={{ lineHeight: '1.6' }}>
              <li><strong>Rule Name:</strong> Hide SharePoint Command Bar</li>
              <li><strong>CSS Selector:</strong> <code>.ms-CommandBar</code> (or an ID like <code>#my-id</code>)</li>
              <li><strong>Action:</strong> Hide</li>
              <li><strong>Script/CSS Rules:</strong> <em>(Leave completely blank)</em></li>
            </ul>
          </div>

          <div style={{ marginBottom: '20px', border: '1px solid #edebe9', padding: '15px', borderRadius: '4px', backgroundColor: '#faf9f8' }}>
            <h3 style={{ marginTop: 0, color: '#0078d4' }}>2. Apply Style (Changing appearance)</h3>
            <p>Applies custom CSS properties to the targeted element.</p>
            <ul style={{ lineHeight: '1.6' }}>
              <li><strong>Rule Name:</strong> Make Site Header Blue</li>
              <li><strong>CSS Selector:</strong> <code>.ms-SiteHeader</code></li>
              <li><strong>Action:</strong> Apply Style</li>
              <li><strong>Script/CSS Rules:</strong> <code>background-color: #0078d4; border-bottom: 2px solid #000;</code> <em>(Do not include {`{}`} braces)</em></li>
            </ul>
          </div>

          <div style={{ marginBottom: '20px', border: '1px solid #edebe9', padding: '15px', borderRadius: '4px', backgroundColor: '#faf9f8' }}>
            <h3 style={{ marginTop: 0, color: '#0078d4' }}>3. Apply Script (Running JavaScript)</h3>
            <p>Executes JavaScript on the page. The SPFx context is automatically available as the <code>context</code> variable.</p>
            <ul style={{ lineHeight: '1.6' }}>
              <li><strong>Rule Name:</strong> Welcome Alert</li>
              <li><strong>CSS Selector:</strong> <em>(Leave blank)</em></li>
              <li><strong>Action:</strong> Apply Script</li>
              <li><strong>Script/CSS Rules:</strong><br />
                <pre style={{ background: '#fff', padding: '10px', border: '1px solid #ccc' }}>
                  {`console.log("The SPFx Context is:", context);\nalert("Welcome to our custom SharePoint Portal!");`}
                </pre>
              </li>
            </ul>
          </div>

          <div style={{ marginBottom: '20px', border: '1px solid #edebe9', padding: '15px', borderRadius: '4px', backgroundColor: '#faf9f8' }}>
            <h3 style={{ marginTop: 0, color: '#0078d4' }}>4. Apply Loader (Full-screen loading animation)</h3>
            <p>Creates a full-screen overlay during initial load. Provide the raw HTML and inline CSS. <strong>Only ONE loader can be active at a time.</strong></p>
            <ul style={{ lineHeight: '1.6' }}>
              <li><strong>Rule Name:</strong> Blue Spinning Loader</li>
              <li><strong>CSS Selector:</strong> <em>(Leave blank)</em></li>
              <li><strong>Action:</strong> Apply Loader</li>
              <li><strong>Script/CSS Rules:</strong><br />
                <pre style={{ background: '#fff', padding: '10px', border: '1px solid #ccc', overflowX: 'auto' }}>
                  {`<style>
  .baz-spinner { 
    border: 8px solid #f3f3f3; 
    border-top: 8px solid #0078d4; 
    border-radius: 50%; 
    width: 60px; height: 60px; 
    animation: spin 1s linear infinite; 
  } 
  @keyframes spin { 
    0% { transform: rotate(0deg); } 
    100% { transform: rotate(360deg); } 
  }
</style>
<div class="baz-spinner"></div>`}
                </pre>
              </li>
            </ul>
          </div>
        </div>
      );
    }

    return (
      <div>
        <div style={{ display: 'flex', gap: '10px', marginBottom: '20px', flexWrap: 'wrap' }}>
          <PrimaryButton text="Add New Rule" iconProps={{ iconName: 'Add' }} onClick={() => { this.setState({ isEditing: true }); }} />
          <DefaultButton text="Export Config" iconProps={{ iconName: 'Download' }} onClick={this.handleExport} />
          <DefaultButton text="Import Config" iconProps={{ iconName: 'Upload' }} onClick={this.triggerImport} />
          <DefaultButton text="Help & Examples" iconProps={{ iconName: 'Help' }} onClick={() => this.setState({ isHelpOpen: true })} />
        </div>

        {/* ACTIVE SECTION */}
        <div
          style={{ ...sectionHeaderStyle, color: '#0078d4', borderColor: '#0078d4' as any }}
          onClick={() => { this.setState({ expandedGroups: { active: !expandedGroups.active, inactive: expandedGroups.inactive } }); }}
        >
          <IconButton iconProps={{ iconName: expandedGroups.active ? 'ChevronDown' : 'ChevronRight' }} styles={{ root: { color: '#0078d4' } }} />
          <h2 style={{ margin: 0, fontSize: '18px' }}>Active Rules ({activeRules.length})</h2>
        </div>
        {expandedGroups.active && (
          <div style={{ paddingLeft: '10px' }}>
            {activeRules.length === 0 ? <p>No active rules.</p> : this.renderGroupedRules(activeRules, 'active')}
          </div>
        )}

        {/* INACTIVE SECTION */}
        <div
          style={{ ...sectionHeaderStyle, color: '#a4262c', borderColor: '#a4262c' as any, marginTop: '40px' }}
          onClick={() => { this.setState({ expandedGroups: { active: expandedGroups.active, inactive: !expandedGroups.inactive } }); }}
        >
          <IconButton iconProps={{ iconName: expandedGroups.inactive ? 'ChevronDown' : 'ChevronRight' }} styles={{ root: { color: '#a4262c' } }} />
          <h2 style={{ margin: 0, fontSize: '18px' }}>Inactive Rules ({inactiveRules.length})</h2>
        </div>
        {expandedGroups.inactive && (
          <div style={{ paddingLeft: '10px' }}>
            {inactiveRules.length === 0 ? <p>No inactive rules.</p> : this.renderGroupedRules(inactiveRules, 'inactive')}
          </div>
        )}
      </div>
    );
  }

  // --- EXPANDABLE SUB-GROUPS ---
  private renderGroupedRules(rules: IRule[], prefix: string) {
    // Sort unique types
    const uniqueTypes = rules
      .map((r) => r.ActionType)
      .filter((value, index, arr) => arr.indexOf(value) === index)
      .sort();

    return (
      <div>
        {uniqueTypes.map((type) => {
          const filtered = rules.filter((r) => r.ActionType === type);
          const groupKey = prefix + "_" + type;
          const isExpanded = this.state.expandedSubGroups[groupKey] !== false;

          return (
            <div key={type} style={{ marginTop: '10px', marginLeft: '5px' }}>

              {/* CLICKABLE SUB-HEADER */}
              <div
                style={{
                  display: 'flex', alignItems: 'center', cursor: 'pointer',
                  backgroundColor: '#f3f2f1', padding: '6px', borderRadius: '4px',
                  marginBottom: '5px', userSelect: 'none'
                }}
                onClick={() => {
                  const newMap = { ...this.state.expandedSubGroups };
                  newMap[groupKey] = !isExpanded;
                  this.setState({ expandedSubGroups: newMap });
                }}
              >
                <IconButton
                  iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
                  styles={{ root: { height: 20, width: 20, marginRight: 8 } }}
                />
                <h4 style={{ margin: 0, color: '#333', textTransform: 'uppercase', fontSize: '12px', fontWeight: 600 }}>
                  {type} ({filtered.length})
                </h4>
              </div>

              {/* RULE ITEMS (Shown only if Expanded) */}
              {isExpanded && (
                <div style={{ paddingLeft: '10px' }}>
                  {filtered.map((rule) => {
                    return (
                      <div key={rule.Id} style={{
                        borderBottom: '1px solid #eee', padding: '8px 0',
                        display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                        backgroundColor: '#fff', paddingLeft: '8px'
                      }}>
                        <div style={{ paddingLeft: '5px', flex: 1 }}>
                          <div style={{ fontWeight: 600, color: rule.IsActive ? '#333' : '#999' }}>{rule.Title}</div>
                          {rule.Selector && <div style={{ fontSize: '11px', color: '#888', fontStyle: 'italic' }}>{rule.Selector}</div>}
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center' }}>
                          <div style={{ marginRight: '15px' }} title="Turn On/Off">
                            <Toggle
                              checked={rule.IsActive}
                              onChange={(_, val) => { void this.toggleRuleActive(rule, !!val); }}
                              styles={{ root: { marginBottom: 0 } }}
                            />
                          </div>
                          <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => { this.setState({ isEditing: true, currentRule: rule }); }} />
                          <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => { this.deleteRule(rule.Id as number); }} styles={{ root: { color: '#a4262c' }, rootHovered: { color: 'red' } }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }

  private renderForm() {
    const currentRule = this.state.currentRule;
    
    // UPGRADE: Replaced deprecated 'onChanged' with 'onChange' to match @fluentui/react v8
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
        <TextField
          label="Rule Name"
          value={currentRule.Title}
          onChange={(e, val) => { this.setState({ currentRule: { ...currentRule, Title: val || '' } }); }}
        />
        <TextField
          label="CSS Selector"
          multiline
          rows={3}
          value={currentRule.Selector}
          onChange={(e, val) => { this.setState({ currentRule: { ...currentRule, Selector: val || '' } }); }}
        />
        <Dropdown
          label="Action"
          selectedKey={currentRule.ActionType}
          options={[
            { key: 'Hide', text: 'Hide' },
            { key: 'Style', text: 'Apply Style' },
            { key: 'Script', text: 'Apply Script' },
            { key: 'Loader', text: 'Apply Loader' }
          ]}
          onChange={(e, item?: IDropdownOption) => { this.setState({ currentRule: { ...currentRule, ActionType: item?.key as string } }); }}
        />
        {(currentRule.ActionType === 'Style' || currentRule.ActionType === 'Script' || currentRule.ActionType === 'Loader') && (
          <TextField
            label="Script/CSS Rules"
            multiline
            rows={3}
            value={currentRule.CSSValue}
            onChange={(e, val) => { this.setState({ currentRule: { ...currentRule, CSSValue: val || '' } }); }}
          />
        )}
        <Toggle
          label="Override (!important)"
          checked={currentRule.IsOverride}
          onChange={(e, val) => { this.setState({ currentRule: { ...currentRule, IsOverride: !!val } }); }}
        />
        <Toggle
          label="Active"
          checked={currentRule.IsActive}
          onChange={(e, val) => { this.setState({ currentRule: { ...currentRule, IsActive: !!val } }); }}
        />
        <div style={{ marginTop: '20px' }}>
          <PrimaryButton text="Save" onClick={this.saveRule} style={{ marginRight: '10px' }} />
          <DefaultButton text="Back" onClick={() => { this.setState({ isEditing: false, currentRule: this.getEmptyRule() }); }} />
        </div>
      </div>
    );
  }
}