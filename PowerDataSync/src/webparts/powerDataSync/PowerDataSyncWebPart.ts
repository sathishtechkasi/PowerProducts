import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption 
} from '@microsoft/sp-webpart-base';

// UPGRADE: Modern SweetAlert import to prevent Webpack chunking errors in SPFx 1.22
import Swal from 'sweetalert2'; 

// UPGRADE: PnP v4 Imports
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import "@pnp/sp/security";
import "@pnp/sp/site-users/web";

import PowerApp, { PowerAppProps } from './components/PowerApp';
import { LoggerService } from './components/LoggerService';
import { LogViewer } from './components/LogViewer'; 

export interface IPowerDataSyncWebPartProps {
  // Setup Status
  isSetupConfirmed: boolean;
  // Logging Configuration
  installType: 'new' | 'existing';
  logListTitle: string;
  logRoleName: string;
  // Metrics Configuration
  metricsListTitle: string;
  metricsLibTitle: string;
  // Versioning
  version?: string;
  enableLogging: boolean;
  showUserAlerts: boolean;
  showHiddenLists: boolean;
}

// UPGRADE: Fixed the corrupted class definition name from the original file
export default class PowerDataSyncWebPart extends BaseClientSideWebPart<IPowerDataSyncWebPartProps> {
  private showLogViewer: boolean = false; 
  private listsDropdownOptions: IPropertyPaneDropdownOption[] = []; 
  
  // UPGRADE: Store the PnP v4 Factory context globally for the web part
  private sp!: SPFI; 

  public async onInit(): Promise<void> {
    return super.onInit().then(async _ => {
      const style = document.createElement('style');
      style.type = 'text/css';
      style.innerHTML = `
    /* GENERIC SELECTOR: Targets any class starting with 'propertyPanePageDescription'  */
    div[class^="propertyPanePageDescription_"], 
    div[class*=" propertyPanePageDescription_"] {
      font-weight: bold !important;
      font-size: x-large !important;
      color: #323130 !important;
      margin-bottom: 15px !important;
      display: block !important;
    }
  `;
      document.head.appendChild(style);
      
      // UPGRADE: Initialize PnP v4
      this.sp = spfi().using(SPFx(this.context));
      
      // 2. Initialize Logger (only if configured)
      LoggerService.init(
        this.context,
        this.properties.metricsListTitle || "PowerDataSyncWebPart",
        this.properties.logListTitle,
        this.properties.enableLogging,
        this.properties.showUserAlerts
      );
      
      // 3. Pre-load lists for the "Use Existing" dropdown
      await this.loadLists();
    });
  }

  // Fetch lists for the Property Pane Dropdown
  private async loadLists(): Promise<void> {
    try {
      // UPGRADE: Replaced legacy SPHttpClient call with PnP v4 query
      const lists = await this.sp.web.lists.filter("Hidden eq false").select("Title", "Id")();
      this.listsDropdownOptions = lists.map((list: any) => {
        return { key: list.Title, text: list.Title };
      });
      // Refresh to show data if pane is already open
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Error loading lists', error);
    }
  }

  public render(): void {
    // -----------------------------------------------------------
    // SCENARIO 1: SHOW WELCOME SCREEN (If not configured)
    // -----------------------------------------------------------
    if (!this.properties.isSetupConfirmed) {
      // Check if user is a Site Admin using legacy context
      const isSiteAdmin = this.context.pageContext.legacyPageContext['isSiteAdmin'];

      // A. NON-ADMIN VIEW (Read Only Message)
      if (!isSiteAdmin) {
        const nonAdminElement = React.createElement('div', {
          style: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            minHeight: '300px',
            padding: '40px',
            textAlign: 'center',
            backgroundColor: '#fff4f4', // Light red background
            borderRadius: '8px',
            border: '1px solid #ffcccc',
            boxShadow: '0 2px 10px rgba(0,0,0,0.05)',
            margin: '20px'
          }
        },
          React.createElement('div', { style: { fontSize: '48px', marginBottom: '15px' } }, '🔒'),
          React.createElement('h2', { style: { margin: '0 0 10px 0', color: '#d13438', fontWeight: 600 } }, 'Setup Required'),
          React.createElement('p', { style: { margin: '0 0 10px 0', color: '#333', maxWidth: '500px', lineHeight: '1.6', fontSize: '15px' } },
            'This web part requires initial configuration and list creation.'
          ),
          React.createElement('p', { style: { margin: '0', color: '#666', maxWidth: '500px', fontSize: '14px', fontStyle: 'italic' } },
            'Please contact your Site Administrator to add and configure this web part.'
          )
        );
        ReactDom.render(nonAdminElement, this.domElement);
        return;
      }

      // B. ADMIN VIEW (Configuration Button)
      const welcomeElement = React.createElement('div', {
        style: {
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          minHeight: '300px',
          padding: '40px',
          textAlign: 'center',
          backgroundColor: '#ffffff',
          borderRadius: '8px',
          boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
          margin: '20px'
        }
      },
        React.createElement('div', { style: { fontSize: '64px', marginBottom: '20px' } }, '⚙️'),
        React.createElement('h2', { style: { margin: '0 0 10px 0', color: '#333', fontWeight: 600 } }, 'Welcome to Power Data Synchronizer'),
        React.createElement('p', { style: { margin: '0 0 20px 0', color: '#666', maxWidth: '500px', lineHeight: '1.6', fontSize: '15px' } },
          'To enjoy the features, please configure the web part first.'
        ),
        React.createElement('div', { style: { fontSize: '13px', color: '#888', marginTop: '10px', padding: '10px', background: '#f5f5f5', borderRadius: '4px' } },
          'For queries, contact ',
          React.createElement('a', { href: 'mailto:admin@Powertechnologies.ae', style: { color: '#0078d4', textDecoration: 'none', fontWeight: 600 } }, 'admin@Powertechnologies.ae')
        ),
        React.createElement('button', {
          style: {
            marginTop: '30px',
            padding: '12px 32px',
            backgroundColor: '#0078d4',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer',
            fontWeight: 600,
            fontSize: '14px',
            boxShadow: '0 4px 6px rgba(0,120,212,0.3)',
            transition: 'background 0.2s'
          },
          onClick: () => { this.context.propertyPane.open(); }
        }, 'Configure Now')
      );
      ReactDom.render(welcomeElement, this.domElement);
      return;
    }

    // -----------------------------------------------------------
    // SCENARIO 2: RENDER APP + LOG VIEWER (If configured)
    // -----------------------------------------------------------

    // 1. Calculate Safe List/Library Names (Prevents 404/500 Errors)
    const metricsList = this.properties.metricsListTitle || "ImportJobs";
    const metricsLib = this.properties.metricsLibTitle || (metricsList + "Documents");

    // 2. Create App Element (PowerApp)
    const appElement = React.createElement(PowerApp as any, {
      context: this.context,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      version: this.properties.version || '1.0.0',
      metricsListTitle: metricsList,
      metricsLibTitle: metricsLib,
      showUserAlerts: this.properties.showUserAlerts,
      showHiddenLists: this.properties.showHiddenLists
    });

    // 3. Log Viewer (Hidden by default, toggled via state)
    const logViewerElement = React.createElement(LogViewer, {
      isOpen: this.showLogViewer,
      context: this.context,
      currentListTitle: "PowerDataSyncWebPart",
      onDismiss: () => {
        this.showLogViewer = false;
        this.render(); // Re-render to close
      }
    });

    // 4. Render container with both components
    ReactDom.render(
      React.createElement('div', null, appElement, logViewerElement),
      this.domElement
    );
  }

  // ==============================================================================
  // PROVISIONING LOGIC
  // ==============================================================================
  private async provisionLogEnvironment(): Promise<void> {
    const { logListTitle, logRoleName } = this.properties;
    if (!logListTitle || !logRoleName) {
      Swal.fire({ icon: 'info', title: 'Missing Info', text: "Please enter both a Log List Name and a Role Name." });
      return;
    }

    try {
      // --- SHOW LOADING POPUP ---
      Swal.fire({
        title: 'Provisioning Logs...',
        html: `Creating list <b>${logListTitle}</b> and configuring permissions. Please wait...`,
        allowOutsideClick: false,
        didOpen: () => { Swal.showLoading(); }
      });

      // UPGRADE: Check if list exists using v4 execution ()
      try {
        await this.sp.web.lists.getByTitle(logListTitle)();
        Swal.fire({ icon: 'info', title: 'Exists', text: `List "${logListTitle}" already exists.` });
      } catch (e) {
        // 1. Create List
        await this.sp.web.lists.add(logListTitle, "System Logs", 100);
        const list = this.sp.web.lists.getByTitle(logListTitle);

        // 2. Add Fields
        await list.fields.addText("Page");
        await list.fields.addText("ItemId");
        await list.fields.addText("Module");
        await list.fields.addText("Severity");
        await list.fields.addMultilineText("Error");
        await list.fields.addText("ErrorId");

        // 3. Add Fields to Default View
        try {
          const view = list.views.getByTitle("All Items");
          const fieldsToShow = ["Page", "ItemId", "Module", "Severity", "Error", "ErrorId"];
          for (const field of fieldsToShow) {
            try { await view.fields.add(field); } catch (viewErr) { }
          }
        } catch (viewFetchErr) {
          console.error("Could not fetch 'All Items' view to update columns.", viewFetchErr);
        }

        // 4. Permissions
        await list.breakRoleInheritance(false, true);
        try {
          const roleDef = await this.sp.web.roleDefinitions.getByName(logRoleName)();
          const everyoneIdentifier = "c:0(.s|true";
          const everyoneUser = await this.sp.web.ensureUser(everyoneIdentifier);
          await list.roleAssignments.add(everyoneUser.Id, roleDef.Id);
        } catch (permErr) {
          console.warn("Could not set permission automatically", permErr);
        }

        Swal.fire({ icon: 'success', title: 'Success', text: `Log List "${logListTitle}" created successfully.` });
      }
    } catch (e: any) {
      console.error("Critical Error in provisionLogEnvironment:", e);
      Swal.fire({ icon: 'error', title: 'Error', text: e.message });
    }
  }

  // --- PROPERTY PANE HELPER METHODS ---
  private getConfigurationActionGroup() {
    return {
      groupName: 'Panel Actions',
      groupFields: [
        PropertyPaneButton('btnSaveConfig', {
          text: 'Save Configuration',
          buttonType: PropertyPaneButtonType.Primary,
          icon: 'Save',
          onClick: () => {
            Swal.fire({ icon: 'success', title: 'Success', text: "Configuration saved successfully!" });
            (this.context.propertyPane as any).close();
          }
        }) as any,
        PropertyPaneButton('btnCancelConfig', {
          text: 'Cancel & Reload',
          buttonType: PropertyPaneButtonType.Normal,
          icon: 'Cancel',
          onClick: () => {
            if (confirm('Are you sure? This will reload the page and revert unsaved UI changes.')) {
              window.location.reload(); // Refresh the page
            }
          }
        }) as any
      ]
    };
  }

  private async provisionMetricsEnvironment(): Promise<void> {
    const listTitle = this.properties.metricsListTitle;
    if (!listTitle) {
      Swal.fire({ icon: 'info', title: 'Missing Info', text: "Please enter a name for the Metrics List." });
      return;
    }
    const libTitle = `${listTitle}Documents`;
    this.properties.metricsLibTitle = libTitle;

    try {
      // --- SHOW LOADING POPUP ---
      Swal.fire({
        title: 'Provisioning Metrics...',
        html: `Creating list <b>${listTitle}</b> and library <b>${libTitle}</b>. Please wait...`,
        allowOutsideClick: false,
        didOpen: () => { Swal.showLoading(); }
      });

      // --- A. Create Metrics List ---
      try {
        await this.sp.web.lists.getByTitle(listTitle)();
      } catch (e) {
        // 1. Create List
        await this.sp.web.lists.add(listTitle, "Stores job history for Power Synchronizer", 100, false);
        const list = this.sp.web.lists.getByTitle(listTitle);

        // 2. Add Columns to List Schema
        await list.fields.addText("DataSycnList");
        await list.fields.addText("OperationType");
        await list.fields.addText("Status");
        await list.fields.addDateTime("JobStartTime");
        await list.fields.addDateTime("JobEndTime");
        await list.fields.addNumber("ItemstoImport");
        await list.fields.addNumber("SuccessCount");
        await list.fields.addNumber("FailureCount");

        // UPGRADE: PnP v4 changed addUrl signature; positional parameters (like `1`) are no longer accepted, we must pass the properties object
        await list.fields.addUrl("SourceFile", { DisplayFormat: 1 } as any);
        await list.fields.addUrl("SuccessFile", { DisplayFormat: 1 } as any);
        await list.fields.addUrl("FailureFile", { DisplayFormat: 1 } as any);

        // 3. Add Columns to Default View
        try {
          const view = list.views.getByTitle("All Items");
          const fieldsToShow = [
            "DataSycnList", "OperationType", "Status",
            "JobStartTime", "JobEndTime", "ItemstoImport",
            "SuccessCount", "FailureCount",
            "SourceFile", "SuccessFile", "FailureFile"
          ];
          for (const fieldName of fieldsToShow) {
            try { await view.fields.add(fieldName); } catch (viewErr) { }
          }
        } catch (viewFetchErr) {
          console.error("Could not fetch 'All Items' view.", viewFetchErr);
        }
      }

      // --- B. Create Documents Library ---
      try {
        await this.sp.web.lists.getByTitle(libTitle)();
      } catch (e) {
        await this.sp.web.lists.add(libTitle, "Stores source and result files", 101, false);
      }

      Swal.fire({
        icon: 'success',
        title: 'Environment Ready',
        text: `Metrics list "${listTitle}" and Library "${libTitle}" have been created and configured.`
      });
    } catch (e: any) {
      Swal.fire({ icon: 'error', title: 'Error', text: e.message });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isSetup = this.properties.isSetupConfirmed;
    // PAGE 1: SETUP WIZARD
    if (!isSetup) {
      return {
        pages: [
          {
            header: { description: 'First Time Setup' },
            groups: [
              {
                groupName: 'Step 1: Logging',
                groupFields: [
                  PropertyPaneChoiceGroup('installType', {
                    label: 'Log Installation',
                    options: [
                      { key: 'new', text: 'Create New Log List', iconProps: { officeFabricIconFontName: 'Add' } },
                      { key: 'existing', text: 'Use Existing', iconProps: { officeFabricIconFontName: 'Link' } }
                    ]
                  }),
                  // CONDITIONAL FIELD: Show TextField OR Dropdown based on choice
                  this.properties.installType === 'new'
                    ? PropertyPaneTextField('logListTitle', {
                      label: 'New Log List Name'
                    })
                    : PropertyPaneDropdown('logListTitle', {
                      label: 'Select Existing Log List',
                      options: this.listsDropdownOptions,
                      selectedKey: this.properties.logListTitle
                    }),
                  PropertyPaneTextField('logRoleName', {
                    label: 'Role Name',
                    value: this.properties.logRoleName,
                    disabled: this.properties.installType === 'existing'
                  }),
                  PropertyPaneButton('btnLogProv', {
                    text: 'Create Log List',
                    buttonType: PropertyPaneButtonType.Primary,
                    onClick: this.provisionLogEnvironment.bind(this),
                    disabled: this.properties.installType === 'existing'
                  })
                ]
              },
              {
                groupName: 'Step 2: Metrics Storage',
                groupFields: [
                  PropertyPaneLabel('lblMetrics', { text: 'Define the list where job history and KPIs will be stored.' }),
                  PropertyPaneTextField('metricsListTitle', {
                    label: 'Metrics List Name',
                    placeholder: 'e.g. ImportJobs',
                    description: 'A document library named [ListName]Documents will also be created.'
                  }),
                  PropertyPaneButton('btnMetricsProv', {
                    text: 'Create Metrics List & Library',
                    buttonType: PropertyPaneButtonType.Primary,
                    onClick: this.provisionMetricsEnvironment.bind(this)
                  })
                ]
              },
              {
                groupName: 'Step 3: Confirmation',
                groupFields: [
                  PropertyPaneToggle('isSetupConfirmed', {
                    label: 'I have created the environment',
                    onText: 'Ready',
                    offText: 'Pending'
                  })
                ]
              }
            ]
          }
        ]
      };
    }
    // PAGE 2: GENERAL SETTINGS (Once configured)
    return {
      pages: [
        {
          header: { description: 'General Configuration' },
          groups: [
            {
              groupName: 'Environment',
              groupFields: [
                PropertyPaneLabel('lblInfo', { text: `Connected to: ${this.properties.metricsListTitle}` }),
                PropertyPaneToggle('showHiddenLists', {
                  label: 'Show Hidden System Lists in Mapping',
                  onText: 'Show All (Template 100)',
                  offText: 'Standard Only',
                  checked: this.properties.showHiddenLists
                }),
                PropertyPaneToggle('isSetupConfirmed', {
                  label: 'Configuration Status',
                  onText: 'Active',
                  offText: 'Disabled'
                })
              ]
            },
            // NEW GROUP: SYSTEM LOGS
            {
              groupName: 'System Logs',
              groupFields: [
                PropertyPaneToggle('enableLogging', {
                  label: 'Enable System Logging',
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('showUserAlerts', {
                  label: 'Show Errors to Users (Popup)',
                  onText: 'Show', offText: 'Hide',
                  checked: this.properties.showUserAlerts,
                  disabled: !this.properties.enableLogging
                }),
                PropertyPaneLabel('lblLogs', {
                  text: 'View system errors and logs captured by this web part.'
                }),
                PropertyPaneButton('btnLogs', {
                  text: 'Open Log Viewer',
                  buttonType: PropertyPaneButtonType.Hero,
                  icon: 'ComplianceAudit',
                  onClick: () => {
                    this.showLogViewer = true;
                    this.render();
                  },
                  disabled: !this.properties.enableLogging
                })
              ]
            }, (this as any).getConfigurationActionGroup()
          ]
        }
      ]
    };
  }
}