import * as React from 'react';
import * as XLSX from 'xlsx';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";
import "@pnp/sp/security";
import { PermissionKind } from "@pnp/sp/security";
import { CommonService } from '../../../Common/Services/CommonService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Swal from 'sweetalert2';
import styles from './PowerDataSync.module.scss';
import { IPowerDataSyncProps, IValidationIssue, IFieldInfo, IExistingJob } from './IPowerDataSyncProps';
import { IPowerDataSyncState } from './IPowerDataSyncState';
import { LoggerService } from './LoggerService';
const systemFields = ['Attachments', 'ContentType', 'Created', 'Modified', 'Author', 'Editor'];
const stripBraces = (g: string) => g.replace(/[{}]/g, "");
const escapeOdataValue = (v: string) => String(v).replace(/'/g, "''");
const norm = (s: string) => (s || '').toLowerCase().replace(/\s+/g, '').replace(/[_-]+/g, '');
// Normalize: trim + convert NBSPs to space + collapse inner spaces
const clean = (s: any): string => {
  try {
    if (s === null || s === undefined) return '';
    return String(s)
      .replace(/\u00A0/g, ' ')     // NBSP -> space
      .replace(/\s+/g, ' ')        // collapse whitespace
      .trim();
  } catch (e: any) {
    void LoggerService.log('System', 'clean', 'Medium', 'N/A', e.message);
    return '';
  }
};
// Tokenize a cell for Choice fields.
// - For MultiChoice: prefer `;#` or `;`. Use `,` only if allowed choices don't contain commas.
// - For Single Choice: DO NOT split on comma; treat cell as a single value.
const splitChoiceCell = (
  raw: any,
  isMulti: boolean,
  allowedList: string[] | undefined
): string[] => {
  try {
    const val = clean(raw);
    if (!val) return [];
    if (!isMulti) {
      // single choice: one token only
      return [val];
    }
    const allowedHasComma = (allowedList || []).some(x => x.indexOf(',') >= 0);
    if (val.indexOf(';#') >= 0) {
      return val.split(/;#/).map(clean).filter(Boolean);
    }
    if (val.indexOf(';') >= 0) {
      return val.split(';').map(clean).filter(Boolean);
    }
    // only consider comma-split if allowed choices themselves don't contain commas
    if (!allowedHasComma && val.indexOf(',') >= 0) {
      return val.split(',').map(clean).filter(Boolean);
    }
    return [val];
  } catch (e: any) {
    void LoggerService.log('System', 'splitChoiceCell', 'Medium', 'N/A', e.message);
    return [];
  }
};
// Unwrap SharePoint "Choices" which may be either string[] or { results: string[] }
const toChoicesArray = (choices: any): string[] => {
  try {
    if (!choices) return [];
    if (Array.isArray(choices)) return choices;
    if (choices && Array.isArray(choices.results)) return choices.results;
    return [];
  } catch (e: any) {
    void LoggerService.log('System', 'toChoicesArray', 'Medium', 'N/A', e.message);
    return [];
  }
};
// Main PowerDataSync component
export default class PowerDataSync extends React.Component<IPowerDataSyncProps, IPowerDataSyncState> {
  private service!: CommonService;
  private workbook: XLSX.WorkBook | null = null;
  private jobStart: Date | null = null;
  private sp!: SPFI;
  constructor(props: IPowerDataSyncProps) {
    super(props);
    try {
      this.service = new CommonService(props.siteUrl, props.context);
      this.sp = spfi().using(SPFx(props.context));
      // Initialize state
      this.state = {
        // Lists
        lists: [],
        selectedList: '',
        workbook: null,
        // Wizard
        currentStep: 1,
        // Excel
        excelData: [],
        columns: [],
        sheetOptions: [],
        selectedSheet: '',
        fileError: '',
        // SP fields
        fields: [],
        excelMapByField: {},
        excelMapIndexByField: {},
        selectedByField: {},
        // Mode
        mode: 'add',
        primaryField: '',
        validationError: '',
        // Progress
        progress: 0,
        completed: 0,
        total: 0,
        isLoading: false,
        // Results
        successUrl: '',
        failureUrl: '',
        successCount: 0,
        failureCount: 0,
        // Row range
        startRow: 1,
        endRow: 0,
        // Validation
        validationIssues: [],
        validationRan: false,
        issueCounts: { required: 0, number: 0, choice: 0, unique: 0, user: 0, total: 0 },
        showAllIssues: false,
        // Permissions
        canAdd: false,
        canEdit: false,
        // Job & file
        jobName: '',
        jobNameError: '',
        jobNameChecking: false,
        jobMode: 'new',
        existingJobs: [],
        selectedJobId: null,
        sourceFileName: '',
        sourceFileBuffer: null,
        sourceFileUrl: '',
        // History
        historyItemId: null,
        // Resume base counts
        baseSuccessCount: 0,
        baseFailureCount: 0,
        basePlannedCount: 0
      };
    } catch (e: any) {
      void LoggerService.log('System', 'constructor', 'High', 'Init', e.message);
    }
  }
  // Load SharePoint lists on mount
  public async componentDidMount(): Promise<void> {
    try {
      const filterParts = ["BaseTemplate eq 100"];
      if (!this.props.showHiddenLists) {
        filterParts.push("Hidden eq false");
      }
      const filterQuery = filterParts.join(" and ");

      const rawLists = await this.sp.web.lists
        .filter(filterQuery) // Apply the dynamic filter
        .select("Title")
        ();

      this.setState({ lists: rawLists.map(function (l) { return l.Title; }) });
      // If a resumeJobId is passed, load that job immediately
      if (this.props.resumeJobId) {
        void this.handleSelectExistingJob(String(this.props.resumeJobId));
      }
    } catch (error: any) {
      console.error('Error loading lists:', error);
      void LoggerService.log('System', 'componentDidMount - ' + (this.state ? this.state.jobName : ''), 'High', 'Init', error.message);
      this.setState({ fileError: 'Failed to load SharePoint lists.' });
    }
  }
  private canGoNext = (): boolean => {
    try {
      const s = this.state;
      if (s.isLoading) return false;
      switch (s.currentStep) {
        case 1:
          return !!s.jobName.trim() &&
            !s.jobNameError &&
            !s.jobNameChecking &&
            !!s.selectedSheet &&
            s.columns.length > 0;
        case 2:
          return !!s.selectedList &&
            s.fields.length > 0 &&
            !s.fields.some(f => f.Required && (!s.excelMapByField[f.InternalName] || !s.selectedByField[f.InternalName]));
        case 3:
          return true; // permissions optional
        case 4:
          return s.validationRan && s.validationIssues.length === 0;
        default:
          return false;
      }
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'canGoNext - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
      return false;
    }
  }
  private setDataSyncState = (updates: Partial<IPowerDataSyncState>, callback?: () => void) => {
    try {
      this.setState(updates as any, callback);
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'setDataSyncState - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  // ---------------- Wizard ----------------
  private goToStep = (n: number) => {
    try {
      if (n < 1) n = 1;
      if (n > 5) n = 5;
      this.setState({ currentStep: n, validationError: '' });
      // If entering step 5 and resume mode, fetch existing jobs (lazy)
      if (n === 5 && this.state.jobMode === 'resume') {
        void this.loadExistingJobs();
      }
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'goToStep - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  private jobNameTimeout: any = null;
  private checkJobName = (name: string) => {
    try {
      const trimmed = (name || '').trim();
      if (!trimmed) {
        this.setState({ jobNameError: 'Job name is required', jobNameChecking: false });
        return;
      }

      // STOP if Metrics List is not configured
      if (!this.props.metricsListTitle) {
        this.setState({ jobNameChecking: false, jobNameError: 'Error: Metrics List not configured.' });
        return;
      }

      if (this.jobNameTimeout) clearTimeout(this.jobNameTimeout);
      this.setState({ jobNameChecking: true, jobNameError: '' });

      this.jobNameTimeout = setTimeout(async () => {
        try {
          const exists = await this.sp.web.lists
            .getByTitle(this.props.metricsListTitle)
            .items.filter(`Title eq '${escapeOdataValue(trimmed)}'`)
            .select('Id')
            .top(1)
            ();

          this.setState({
            jobNameError: exists.length > 0 ? 'Job name already exists' : '',
            jobNameChecking: false
          });
        } catch (err: any) {
          console.error('Job name check failed:', err);
          // IF 404, IT MEANS THE LIST DOES NOT EXIST. DO NOT ALLOW PROCEEDING.
          if (err.message && err.message.indexOf('404') !== -1) {
            this.setState({
              jobNameError: `Error: List '${this.props.metricsListTitle}' not found. Please configure the web part.`,
              jobNameChecking: false
            });
          } else {
            this.setState({ jobNameError: '', jobNameChecking: false });
          }
        }
      }, 400);
    } catch (e: any) {
      void LoggerService.log('System', 'checkJobName', 'Medium', 'N/A', e.message);
    }
  }
  private fetchSourceFile = async (url: string): Promise<Uint8Array> => {
    try {
      const serverRelativeUrl = url.replace(window.location.origin, '');
      const file = this.sp.web.getFileByServerRelativePath(serverRelativeUrl);
      const buffer = await file.getBuffer();
      return new Uint8Array(buffer);
    } catch (err: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'fetchSourceFile - ' + this.state.jobName, 'High', this.state.mode, err.message);
      throw new Error(`Failed to load source file: ${err && (err as any).message ? (err as any).message : String(err)}`);
    }
  }
  private validateStep = (step: number): boolean => {
    try {
      const {
        selectedSheet, selectedList, columns, fields,
        excelMapByField, selectedByField, mode, primaryField, validationIssues, validationRan
      } = this.state;
      this.setState({ validationError: '' });
      if (step === 0) {
        if (!selectedSheet || columns.length === 0) {
          this.setState({ validationError: 'Please upload an Excel file and select a sheet.' });
          return false;
        }
      }
      if (step === 1) {
        if (this.state.jobMode === 'new') {
          const name = (this.state.jobName || '').trim();
          if (!name) {
            this.setState({ validationError: 'Please enter a Job Name before continuing.' });
            return false;
          }
        } else {
          if (!this.state.selectedJobId) {
            this.setState({ validationError: 'Please select a failed job to resume.' });
            return false;
          }
        }
        if (!selectedSheet || columns.length === 0) {
          this.setState({ validationError: 'Please upload an Excel file and select a sheet.' });
          return false;
        }
      }
      if (step === 2) {
        if (!selectedList) {
          this.setState({ validationError: 'Please select a SharePoint list.' });
          return false;
        }
        const missingRequired: string[] = [];
        for (let i = 0; i < fields.length; i++) {
          const f = fields[i];
          if (f.Required) {
            const mapped = !!excelMapByField[f.InternalName];
            const checked = !!selectedByField[f.InternalName];
            if (!mapped || !checked) missingRequired.push(f.Title || f.InternalName);
          }
        }
        if (missingRequired.length) {
          this.setState({
            validationError: 'Please map and check all required fields: ' + missingRequired.join(', ')
          });
          return false;
        }
      }
      if (step === 3) {
        if (mode === 'update' && !primaryField) {
          this.setState({ validationError: 'Select a primary key (Excel column) for Update mode.' });
          /* return false;*/
        }
        if (mode === 'add' && !this.state.canAdd) {
          this.setState({ validationError: `You don't have permission to add items to "${this.state.selectedList}".` });
          /* return false;*/
        }
        if (mode === 'update' && !this.state.canEdit) {
          this.setState({ validationError: `You don't have permission to edit items in "${this.state.selectedList}".` });
          /* return false;*/
        }
      }
      if (step === 4) {
        if (!validationRan || (validationIssues && validationIssues.length > 0)) {
          this.setState({ validationError: 'Please fix the data validation errors before continuing.' });
          return false;
        }
      }
      return true;
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'validateStep - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
      return false;
    }
  }
  private handleBack = () => {
    try {
      this.goToStep(this.state.currentStep - 1);
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'handleBack - ' + this.state.jobName, 'Low', this.state.mode, e.message);
    }
  }
  private handleNext = async () => {
    try {
      const currentStep = this.state.currentStep;
      if (currentStep === 3) {
        await this.runDataValidation();
        this.goToStep(4);
        return;
      }
      if (this.validateStep(currentStep)) {
        this.goToStep(currentStep + 1);
      }
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'handleNext - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  // ----------- Job (resume) helpers -----------
  private loadExistingJobs = async () => {
    try {
      if (this.state.existingJobs.length) return;
      this.setState({ isLoading: true });
      try {
        const jobs = await this.sp.web.lists
          .getByTitle(this.props.metricsListTitle)
          .items.select('Id', 'Title', 'Status')
          .filter("Status eq 'Failed' or Status eq 'Completed with errors'")
          .top(100)
          ();
        this.setState({ existingJobs: jobs });
      } catch (err: any) {
        console.error(err);
        void LoggerService.log(this.props.metricsListTitle, 'loadExistingJobs - ' + this.state.jobName, 'High', this.state.mode, err.message);
      } finally {
        this.setState({ isLoading: false });
      }
    } catch (e: any) {
      void LoggerService.log('System', 'loadExistingJobs(wrapper) - ' + this.state.jobName, 'High', this.state.mode, e.message);
    }
  }
  //  -------------- Job mode --------------
  private handleJobModeChange = (mode: 'new' | 'resume') => {
    try {
      this.setState({
        jobMode: mode,
        validationError: '',
        selectedJobId: null,
        jobName: mode === 'new' ? '' : this.state.jobName
      }, () => {
        if (mode === 'resume') this.loadExistingJobs();
      });
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'handleJobModeChange - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  private handleSelectExistingJob = async (idStr: string) => {
    try {
      const id = parseInt(idStr, 10);
      if (isNaN(id)) {
        this.setState({ selectedJobId: null, jobName: '' });
        return;
      }
      // ---- 1. UI title ----
      const job = this.state.existingJobs.filter(j => j.Id === id)[0];
      // If job not in list (e.g. direct resume), we proceed blindly and let fetch fail if invalid
      if (!job && !this.props.resumeJobId && !idStr) return;
      this.setState({
        selectedJobId: id,
        jobName: job ? job.Title : 'Resuming Job...',
        isLoading: true,
        validationError: '',
        excelData: [], columns: [], sheetOptions: [], selectedSheet: '',
        sourceFileBuffer: null, sourceFileUrl: '', workbook: null
      });
      try {
        // ---- 2. DataSycnHistory item ----
        const historyItem: any = await this.sp.web.lists.getByTitle(this.props.metricsListTitle).items
          .getById(id)
          .select('Id,Title,DataSycnList,OperationType,SourceFile,FailureFile,SuccessCount,FailureCount,ItemstoImport')
          .expand('File')
          ();
        const srcFile = historyItem.SourceFile && historyItem.SourceFile.Url ? historyItem.SourceFile.Url : '';
        const failFile = historyItem.FailureFile && historyItem.FailureFile.Url ? historyItem.FailureFile.Url : '';
        // ---- 3. Download source file (PnP-JS v1) ----
        let buffer: Uint8Array | null = null;
        let srcUrl = '';
        // Priority: Try Failure file first (to resume errors), else Source file
        // NOTE: For a strict resume of "failed items", logic might need adjustment.
        // Here we assume user wants to fix the source or re-run.
        // Actually, usually "Resume" implies retrying the source or the failed rows.
        // If we pick failure file, we get only failed rows.
        if (failFile) {
          srcUrl = failFile;
        } else if (srcFile) {
          srcUrl = srcFile;
        } else {
          throw new Error('Source file missing in DataSycnHistory.');
        }
        try {
          const serverRelativeUrl = srcUrl.replace(window.location.origin, '');
          const file = this.sp.web.getFileByServerRelativePath(serverRelativeUrl);
          const arrayBuffer: ArrayBuffer = await file.getBuffer();
          buffer = new Uint8Array(arrayBuffer);
        } catch (dlErr) {
          throw new Error('Could not download source/failure file – it may have been deleted.');
        }
        // ---- 4. Parse workbook & pick sheet from failure file name ----
        const wb = XLSX.read(buffer, { type: 'array' });
        const sheetNames = wb.SheetNames;
        let targetSheet = sheetNames[0];
        // const failureFileName = failFile && (failFile.Name || failFile.FileName);
        // if (failureFileName) {
        //   const extracted = this.extractSheetFromFailureFileName(failureFileName);
        //   if (extracted && sheetNames.indexOf(extracted) !== -1) {
        //     targetSheet = extracted;
        //   }
        // }
        // ---- 5. Load sheet data (Array of Objects) ----
        const ws = wb.Sheets[targetSheet];
        const json: any[] = XLSX.utils.sheet_to_json(ws);
        const cols = json.length
          ? Object.keys(json[0]).filter(function (c) { return c && json.some(function (r) { return r[c] != null && r[c] !== ''; }); })
          : [];
        // ---- 6. Restore full state (same as fresh upload) ----
        this.setState({
          jobMode: 'resume',
          sourceFileName: srcFile.Description || srcFile.FileName || 'Resumed file',
          sourceFileUrl: srcUrl,
          sourceFileBuffer: buffer,
          workbook: wb,
          sheetOptions: sheetNames,
          selectedSheet: targetSheet,
          columns: cols,
          excelData: json,
          selectedList: historyItem.DataSycnList || '',
          mode: (historyItem.OperationType || '').toLowerCase().indexOf('update') !== -1 ? 'update' : 'add',
          startRow: 2,
          endRow: json.length,
          baseSuccessCount: historyItem.SuccessCount || 0,
          baseFailureCount: historyItem.FailureCount || 0,
          basePlannedCount: historyItem.ItemstoImport || 0,
          historyItemId: id,
          currentStep: 2,
          isLoading: false,
          excelMapByField: {},     // ← clean slate
          selectedByField: {}
        }, async () => {
          // Reuse logic to load fields
          if (historyItem.DataSycnList) {
            try {
              // Check perms
              const perms = await this.checkListPermissions(historyItem.DataSycnList);
              const canAdd = perms.canAdd || perms.unknown;
              const canEdit = perms.canEdit || perms.unknown;
              const fieldInfos = await this.sp.web.lists.getByTitle(historyItem.DataSycnList).fields
                .filter("ReadOnlyField eq false and Hidden eq false")
                .select('InternalName', 'Title', 'TypeAsString', 'Required', 'Choices', 'EnforceUniqueValues')
                ();
              const fields: IFieldInfo[] = fieldInfos
                .filter(f => systemFields.indexOf(f.InternalName) === -1 && f.InternalName !== 'ContentType')
                .map(f => ({
                  InternalName: f.InternalName,
                  Title: f.Title,
                  TypeAsString: f.TypeAsString,
                  Required: f.Required,
                  Choices: toChoicesArray(f.Choices),
                  EnforceUniqueValues: !!f.EnforceUniqueValues
                }));
              const selectedByField: { [key: string]: boolean } = {};
              fields.forEach(function (f) { selectedByField[f.InternalName] = true; });
              this.setState({
                fields: fields,
                selectedByField: selectedByField,
                excelMapByField: {},
                canAdd,
                canEdit
              }, () => {
                this.autoMapFields();
                const { excelMapByField, fields: updatedFields } = this.state;
                const newSelected = { ...this.state.selectedByField };
                updatedFields.forEach(f => {
                  if (f.Required && excelMapByField[f.InternalName]) {
                    newSelected[f.InternalName] = true;
                  }
                });
                this.setState({ selectedByField: newSelected });
              });
            } catch (err: any) {
              console.error('Failed to load fields for resume', err);
              void LoggerService.log(historyItem.DataSycnList || 'System', 'handleSelectExistingJob-fields - ' + this.state.jobName, 'High', this.state.mode, err.message);
            }
          }
        });
      } catch (err: any) {
        console.error('Resume failed', err);
        void LoggerService.log('System', 'handleSelectExistingJob - ' + this.state.jobName, 'High', this.state.mode, err.message);
        this.setState({
          isLoading: false,
          validationError: `Resume failed: ${(err as any).message || err}`
        });
      }
    } catch (e: any) {
      void LoggerService.log('System', 'handleSelectExistingJob(wrapper) - ' + this.state.jobName, 'High', this.state.mode, e.message);
    }
  }
  private loadFieldsForResume = (listName: string, callback: () => void) => {
    try {
      this.sp.web.lists.getByTitle(listName).fields
        .filter("ReadOnlyField eq false and Hidden eq false")
        .select('InternalName', 'Title', 'TypeAsString', 'Required', 'Choices', 'EnforceUniqueValues')
        ()
        .then(fieldInfos => {
          const fields: IFieldInfo[] = fieldInfos
            .filter(f => systemFields.indexOf(f.InternalName) === -1)
            .map(f => ({
              InternalName: f.InternalName,
              Title: f.Title,
              TypeAsString: f.TypeAsString,
              Required: f.Required,
              Choices: toChoicesArray(f.Choices),
              EnforceUniqueValues: !!f.EnforceUniqueValues
            }));
          const selectedByField: { [key: string]: boolean } = {};
          fields.forEach(f => { selectedByField[f.InternalName] = true; });
          this.setState({
            fields,
            selectedByField,
            excelMapByField: {}  // clean
          }, callback);  // ← callback runs after fields set
        })
        .catch(err => {
          void LoggerService.log(listName || 'System', 'loadFieldsForResume - ' + this.state.jobName, 'High', this.state.mode, err.message);
        });
    } catch (e: any) {
      void LoggerService.log(listName || 'System', 'loadFieldsForResume(wrapper) - ' + this.state.jobName, 'High', this.state.mode, e.message);
    }
  }
  private loadFileFromUrl = async (url: string): Promise<Uint8Array> => {
    try {
      const serverRelativeUrl = url.replace(window.location.origin, '');
      const file = this.sp.web.getFileByServerRelativePath(serverRelativeUrl);
      const buffer = await file.getBuffer();
      return new Uint8Array(buffer);
    } catch (err: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'loadFileFromUrl - ' + this.state.jobName, 'High', this.state.mode, err.message);
      throw new Error(`Could not download source file: ${err.message}`);
    }
  }
  // -------------- File & Sheet --------------
  private handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    try {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      const allowedTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv'
      ];
      const maxSize = 10485760;
      if (allowedTypes.indexOf(file.type) === -1) {
        this.setState({ fileError: 'Invalid file type.' });
        return;
      }
      if (file.size > maxSize) {
        this.setState({ fileError: 'File exceeds 10MB.' });
        return;
      }
      // Reset prior state
      this.setState({
        fileError: '',
        successUrl: '',
        failureUrl: '',
        successCount: 0,
        failureCount: 0,
        validationError: '',
        startRow: 1,
        endRow: 0,
        validationIssues: [],
        validationRan: false
      });
      const reader = new FileReader();
      reader.onload = (ev: any) => {
        try {
          const data = new Uint8Array(ev.target.result);
          this.workbook = XLSX.read(data, { type: 'array' });
          // For DataSycnHistoryDocuments/JobName/Source
          this.setState({
            sourceFileName: file.name,
            sourceFileBuffer: data
          });
          const names = this.workbook.SheetNames;
          const sheet = this.workbook.Sheets[names[0]];
          const json: any[] = XLSX.utils.sheet_to_json(sheet);
          const cols = json.length
            ? Object.keys(json[0]).filter(function (c) { return c && json.some(function (r) { return r[c] != null && r[c] !== ''; }); })
            : [];
          this.setState({
            sheetOptions: names,
            selectedSheet: names[0],
            excelData: json,
            columns: cols,
            startRow: 1,
            endRow: json.length
          }, this.autoMapFields);
        } catch (err: any) {
          console.error(err);
          void LoggerService.log('System', 'handleFileUpload(reader) - ' + this.state.jobName, 'High', this.state.mode, err.message);
          this.setState({ fileError: 'Error parsing file.' });
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (e: any) {
      void LoggerService.log('System', 'handleFileUpload - ' + this.state.jobName, 'High', this.state.mode, e.message);
    }
  }
  private handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    try {
      const sel = e.target.value;
      if (!this.workbook) return;
      const json: any[] = XLSX.utils.sheet_to_json(this.workbook.Sheets[sel]);
      const cols = json.length
        ? Object.keys(json[0]).filter(function (c) { return c && json.some(function (r) { return r[c] != null && r[c] !== ''; }); })
        : [];
      this.setState({
        selectedSheet: sel,
        excelData: json,
        columns: cols,
        startRow: 1,
        endRow: json.length,
        validationIssues: [],
        validationRan: false
      }, this.autoMapFields);
    } catch (e: any) {
      void LoggerService.log('System', 'handleSheetChange - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  // Safe wrapper for hasPermissions (avoid runtime throws on odd masks)
  private has(perms: any, kind: PermissionKind): boolean {
    try {
      return this.sp.web.hasPermissions(perms, kind) === true;
    } catch (e: any) {
      return false;
    }
  }
  /**
   * Best-effort permission check for a list:
   *  1) current user on LIST
   *  2) explicit user on LIST
   *  3) WEB-level heuristic (marks as "unknown" so we don't block)
   *
   * NOTE: We still do a definitive permission probe at run-time before importing.
   */
  private async checkListPermissions(
    title: string
  ): Promise<{ canAdd: boolean; canEdit: boolean; unknown: boolean }> {
    let canAdd = false;
    let canEdit = false;
    let unknown = false;
    try {
      const list = this.sp.web.lists.getByTitle(title);
      // 1) LIST perms for current user
      try {
        const p1: any = await list.getCurrentUserEffectivePermissions();
        canAdd = this.has(p1, PermissionKind.AddListItems);
        canEdit = this.has(p1, PermissionKind.EditListItems);
      } catch (e1: any) {
        void LoggerService.log(title, 'checkListPermissions(list) - ' + this.state.jobName, 'Low', this.state.mode, e1.message);
      }
      // 2) If both false, try explicit user on LIST
      if (!canAdd && !canEdit) {
        try {
          const me: any = await this.sp.web.currentUser();
          if (me && me.LoginName) {
            const p2: any = await list.getUserEffectivePermissions(me.LoginName);
            canAdd = this.has(p2, PermissionKind.AddListItems);
            canEdit = this.has(p2, PermissionKind.EditListItems);
          }
        } catch (e2: any) {
          void LoggerService.log(title, 'checkListPermissions(user) - ' + this.state.jobName, 'Low', this.state.mode, e2.message);
        }
      }
      // 3) If still false, check WEB perms as heuristic (some SE farms return only web perms)
      if (!canAdd && !canEdit) {
        try {
          const wp: any = await this.sp.web.getCurrentUserEffectivePermissions();
          const webAdd = this.has(wp, PermissionKind.AddListItems);
          const webEdit = this.has(wp, PermissionKind.EditListItems);
          if (webAdd || webEdit) {
            // We can’t prove list-level rights, but web looks OK; don’t block UI.
            unknown = true;
          }
        } catch (e3: any) {
          void LoggerService.log(title, 'checkListPermissions(web) - ' + this.state.jobName, 'Low', this.state.mode, e3.message);
        }
      }
    } catch (outer: any) {
      // If the list lookup itself failed, mark as unknown so UI won’t block
      unknown = true;
      void LoggerService.log(title, 'checkListPermissions(outer) - ' + this.state.jobName, 'Medium', this.state.mode, outer.message);
    }
    return { canAdd, canEdit, unknown };
  }
  // -------------- List & Fields --------------
  private handleListChange = async (e: React.ChangeEvent<HTMLSelectElement>): Promise<void> => {
    try {
      const title = e.target.value;
      this.setState({
        selectedList: title,
        excelMapByField: {},
        selectedByField: {},
        validationError: '',
        validationIssues: [],
        validationRan: false,
        canAdd: false,
        canEdit: false
      });
      if (!title) return;
      try {
        const allFields = await this.sp.web.lists
          .getByTitle(title)
          .fields
          .filter("Hidden eq false and ReadOnlyField eq false")
          .select("Title,InternalName,TypeAsString,LookupList,LookupField,AllowMultipleValues,Required,EnforceUniqueValues,Choices,FillInChoice")
          ();
        const visibleFields = allFields.filter((f: any) => systemFields.indexOf(f.InternalName) === -1);
        // normalize "Choices" to string[] for Choice/MultiChoice fields
        for (let i = 0; i < visibleFields.length; i++) {
          const f: any = visibleFields[i];
          if (/Choice/i.test(f.TypeAsString)) {
            f.Choices = toChoicesArray(f.Choices);
          }
        }
        this.setState({ fields: visibleFields }, this.autoMapFields);
        this.setState({ fields: visibleFields }, this.autoMapFields);
      } catch (error: any) {
        console.error('Error loading fields:', error);
        void LoggerService.log(title || 'System', 'handleListChange-fields - ' + this.state.jobName, 'High', this.state.mode, error.message);
        this.setState({ fileError: 'Failed loading fields.' });
      }
      // Permissions for the selected list
      try {
        const r = await this.checkListPermissions(title);
        this.setState({
          canAdd: r.canAdd || r.unknown,   // <- don’t block if unknown
          canEdit: r.canEdit || r.unknown, // <- don’t block if unknown
        });
        /*TODO*/
        /* this.setState({ canAdd: true, canEdit: true });*/
      } catch (permErr: any) {
        console.warn('Could not read permissions for list', permErr);
        void LoggerService.log(title || 'System', 'handleListChange-perms - ' + this.state.jobName, 'Medium', this.state.mode, permErr.message);
        this.setState({ canAdd: true, canEdit: true });
      }
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'handleListChange(wrapper) - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  // -------------- Auto-map --------------
  //New Mode works fine.// Resume Mode Except automap works
  private autoMapFields = (): void => {
    try {
      const fields = this.state.fields;
      const columns = this.state.columns;
      if (!fields.length || !columns.length) return;
      const byField: { [k: string]: string } = {};
      const selected: { [k: string]: boolean } = {};
      const normCols = columns.map(c => ({ raw: c, n: norm(c) }));
      const assigned: { [col: string]: boolean } = {};
      for (let i = 0; i < fields.length; i++) {
        const f = fields[i];
        const candidates = [f.Title, f.InternalName].map(norm);
        const matches = normCols.filter(c => candidates.indexOf(c.n) !== -1);
        const hit = matches.length ? matches[0] : null;
        const chosen = (hit && !assigned[hit.raw]) ? hit.raw : '';
        byField[f.InternalName] = chosen;
        selected[f.InternalName] = !!chosen;
        if (chosen) assigned[chosen] = true;
      }
      this.setState({ excelMapByField: byField, selectedByField: selected });
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'autoMapFields - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  private handleExcelMapChange = (internal: string, col: string) => {
    try {
      this.setState(function (prev) {
        const newMap: { [k: string]: string } = {};
        for (const k in prev.excelMapByField) if (Object.prototype.hasOwnProperty.call(prev.excelMapByField, k)) newMap[k] = prev.excelMapByField[k];
        if (col) {
          for (const k in newMap) {
            if (Object.prototype.hasOwnProperty.call(newMap, k) && k !== internal && newMap[k] === col) {
              newMap[k] = '';
            }
          }
        }
        newMap[internal] = col;
        const newSel: { [k: string]: boolean } = {};
        for (const k in prev.selectedByField) if (Object.prototype.hasOwnProperty.call(prev.selectedByField, k)) newSel[k] = prev.selectedByField[k];
        newSel[internal] = !!col || !!prev.selectedByField[internal];
        return {
          excelMapByField: newMap,
          selectedByField: newSel,
          validationIssues: [],
          validationRan: false
        } as any;
      });
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'handleExcelMapChange - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  // Map the chosen Excel primary column to the SP internal name
  private getSpFieldForExcelselection(): string | null {
    try {
      const { excelMapByField, primaryField } = this.state;
      if (!primaryField) return null;
      for (const internal in excelMapByField) {
        if (Object.prototype.hasOwnProperty.call(excelMapByField, internal) &&
          excelMapByField[internal] === primaryField) {
          return internal;
        }
      }
      return null;
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'getSpFieldForExcelselection - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
      return null;
    }
  }
  private handleFieldCheck = (internal: string, checked: boolean) => {
    try {
      this.setState(function (prev) {
        const newSel: { [k: string]: boolean } = {};
        for (const k in prev.selectedByField) if (Object.prototype.hasOwnProperty.call(prev.selectedByField, k)) newSel[k] = prev.selectedByField[k];
        newSel[internal] = checked;
        return { selectedByField: newSel, validationIssues: [], validationRan: false } as any;
      });
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'handleFieldCheck - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  // Ensure Choices (and FillInChoice) are present for Choice/MultiChoice
  private async ensureChoicesForField(listTitle: string, f: IFieldInfo): Promise<IFieldInfo> {
    try {
      if (!(/Choice/i.test(f.TypeAsString))) return f;
      let anyF: any = f;
      let needsChoices = !(anyF.Choices && anyF.Choices.length);
      let lacksFillIn = typeof anyF.FillInChoice === 'undefined';
      if (!needsChoices && !lacksFillIn) return f;
      try {
        let full: any = await this.sp.web.lists.getByTitle(listTitle)
          .fields.getByInternalNameOrTitle(f.InternalName)
          ();
        if (full) {
          if (full.Choices) anyF.Choices = full.Choices as string[];
          if (typeof full.FillInChoice !== 'undefined') anyF.FillInChoice = !!full.FillInChoice;
        }
      } catch (e: any) {
        void LoggerService.log(listTitle, 'ensureChoicesForField(inner) - ' + this.state.jobName, 'Low', this.state.mode, e.message);
      }
      return f;
    } catch (e: any) {
      void LoggerService.log(listTitle, 'ensureChoicesForField(wrapper) - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
      return f;
    }
  }
  // Re-validates the current selection/range and fills state.validationIssues + issueCounts
  private runDataValidation = async (): Promise<void> => {
    try {
      let excelData = this.state.excelData;
      let fields = this.state.fields;
      let excelMapByField = this.state.excelMapByField;
      let selectedByField = this.state.selectedByField;
      let startRow = this.state.startRow || 1;
      let endRow = (this.state.endRow || excelData.length);
      let selectedList = this.state.selectedList;
      const splitTokens = (raw: any): string[] => {
        try {
          if (raw === null || raw === undefined) return [];
          return String(raw)
            .replace(/;#/g, ';')
            .split(/[;,]/)
            .map(function (s) { return s.trim(); })
            .filter(function (s) { return !!s; });
        } catch (e: any) {
          void LoggerService.log('System', 'splitTokens', 'Low', 'N/A', e.message);
          return [];
        }
      };
      // Helper: tokenize multi-value cells like "A;#B", "A;B", or "A,B"
      // Build Choice metadata (Allowed values + FillInChoice) per field
      let choiceAllowedByField: { [internal: string]: { [lower: string]: boolean } } = {};
      let choiceAllowedListByField: { [internal: string]: string[] } = {};
      let fillInChoiceByField: { [internal: string]: boolean } = {};
      // First pass: use what we already have on the field (Choices / FillInChoice)
      for (let i = 0; i < fields.length; i++) {
        let f1 = fields[i];
        if (!(/Choice/i).test(f1.TypeAsString)) continue;
        let allowedList: string[] = toChoicesArray((f1 as any).Choices);
        let fillIn = !!(f1 as any).FillInChoice;
        // If we don't have Choices or FillInChoice yet, attempt to read from the list field definition
        if ((!allowedList || !allowedList.length) || (typeof (f1 as any).FillInChoice === 'undefined')) {
          if (selectedList) {
            try {
              let def: any = await this.sp.web.lists
                .getByTitle(selectedList)
                .fields
                .getByInternalNameOrTitle(f1.InternalName)
                .select('Choices,FillInChoice')
                ();
              if (def && def.Choices && def.Choices.length) {
                allowedList = def.Choices;
              }
              if (def && typeof def.FillInChoice !== 'undefined') {
                fillIn = !!def.FillInChoice;
              }
            } catch (e: any) {
              void LoggerService.log(selectedList, 'runDataValidation(choices) - ' + this.state.jobName, 'Low', this.state.mode, e.message);
            }
          }
        }
        // Normalize allowed choices
        let map: { [lower: string]: boolean } = {};
        let normalized: string[] = [];
        if (allowedList && allowedList.length) {
          for (let a = 0; a < allowedList.length; a++) {
            let ch = clean(allowedList[a]);
            if (ch) {
              map[ch.toLowerCase()] = true;
              normalized.push(ch);
            }
          }
        }
        choiceAllowedByField[f1.InternalName] = map;
        choiceAllowedListByField[f1.InternalName] = normalized;
        fillInChoiceByField[f1.InternalName] = !!fillIn;
      }
      // Track in-sheet uniqueness for fields that enforce unique
      let uniqueTrack: { [internal: string]: { seen: { [val: string]: number }, title: string } } = {};
      for (let u = 0; u < fields.length; u++) {
        let fu = fields[u];
        if ((fu as any).EnforceUniqueValues) {
          uniqueTrack[fu.InternalName] = { seen: {}, title: fu.Title || fu.InternalName };
        }
      }
      // Collect which user values exist to validate (batch ensure)
      let userFields: any[] = [];
      for (let uf = 0; uf < fields.length; uf++) {
        let fU = fields[uf];
        if ((/User/i).test(fU.TypeAsString) && selectedByField[fU.InternalName] && excelMapByField[fU.InternalName]) {
          userFields.push(fU);
        }
      }
      //
      let userValidValueByField: { [internal: string]: { [valLower: string]: boolean } } = {};
      let from = Math.max(1, Math.min(startRow, excelData.length)) - 1;
      let to = Math.max(from, Math.min(endRow, excelData.length)) - 1;
      // Collect distinct user tokens per user field across the selected range
      for (let uf2 = 0; uf2 < userFields.length; uf2++) {
        let f2 = userFields[uf2];
        let internalUser = f2.InternalName;
        let colUser = excelMapByField[internalUser];
        let seenVals: { [valLower: string]: boolean } = {};
        for (let idxi = from; idxi <= to; idxi++) {
          let rowi = excelData[idxi];
          let rawi = rowi[colUser];
          if (rawi == null || String(rawi).trim() === '') continue;
          let isMultiUser = (/Multi/i).test(f2.TypeAsString);
          if (isMultiUser) {
            let parts0 = splitTokens(rawi);
            for (let p0 = 0; p0 < parts0.length; p0++) {
              let v0 = parts0[p0];
              let low0 = v0.toLowerCase();
              if (!seenVals[low0]) seenVals[low0] = false;
            }
          } else {
            let v1 = String(rawi).trim();
            if (v1) {
              let low1 = v1.toLowerCase();
              if (!seenVals[low1]) seenVals[low1] = false;
            }
          }
        }
        userValidValueByField[internalUser] = seenVals;
      }
      // Validate the collected user identities exist (best-effort)
      for (let uf3 = 0; uf3 < userFields.length; uf3++) {
        let f3 = userFields[uf3];
        let internal3 = f3.InternalName;
        let validMap = userValidValueByField[internal3];
        for (let low in validMap) {
          if (!Object.prototype.hasOwnProperty.call(validMap, low)) continue;

          // --- STRICT FORMAT CHECK ---
          // Must start with 'i:0#.w|' (standard claims prefix)
          if (low.indexOf('i:0#.w|') !== 0) {
            // Add an issue immediately if format is wrong
            // We have to find at least one row index that used this invalid value to report it effectively.
            // Since validMap is aggregated, we'll mark it false here, 
            // and the main loop below will catch the specific rows and add the "User not found" error.
            // Ideally, we add a specific "Invalid Format" error here, but marking false ensures it fails validation.
            validMap[low] = false;
            continue;
          }

          try {
            await this.sp.web.ensureUser(low);
            validMap[low] = true;
          } catch (e: any) {
            validMap[low] = false;
          }
        }
      }
      // Now perform row-by-row validation
      let issues: IValidationIssue[] = [];
      for (let idx = from; idx <= to; idx++) {
        let row = excelData[idx];
        // Displayed row number within selected range (1-based)
        let dispRow = (idx - from + 1);
        for (let fi = 0; fi < fields.length; fi++) {
          let f = fields[fi];
          let internal = f.InternalName;
          if (!selectedByField[internal]) continue;
          let excelCol = excelMapByField[internal];
          if (!excelCol) continue;
          let raw = row[excelCol];
          // --- Required check ---
          if (f.Required) {
            if (raw === null || raw === undefined || this.cleanCell(raw) === '') {
              issues.push({ row: dispRow, column: (f.Title || internal), detail: 'Required value is missing.', type: 'required' });
            }
          }
          // --- Number-like fields ---
          if ((/Number|Currency|Counter/i).test(f.TypeAsString)) {
            if (!(raw === null || raw === undefined || this.cleanCell(raw) === '')) {
              let s = this.cleanCell(raw);
              let asNum = parseFloat(s);
              if (isNaN(asNum)) {
                issues.push({ row: dispRow, column: (f.Title || internal), detail: 'Expected a number but found "' + s + '".', type: 'number' });
              }
            }
          }
          // --- Choice / MultiChoice strict validation (unless Fill-in is enabled) ---
          if ((/Choice/i).test(f.TypeAsString)) {
            const fillIn = !!fillInChoiceByField[internal];
            if (!fillIn) {
              const allowedMap = choiceAllowedByField[internal] || {};
              const allowedVals = choiceAllowedListByField[internal] || [];
              const isMultiChoice = (/MultiChoice/i).test(f.TypeAsString) || f.AllowMultipleValues === true;
              if (!(raw === null || raw === undefined || this.cleanCell(raw) === '')) {
                const vals = splitChoiceCell(raw, isMultiChoice, allowedVals);
                for (let vm = 0; vm < vals.length; vm++) {
                  const vv = clean(vals[vm]);
                  if (vv && !allowedMap[vv.toLowerCase()]) {
                    issues.push({
                      row: dispRow,
                      column: (f.Title || internal),
                      detail: 'Invalid choice "' + vv + '". Allowed: ' + allowedVals.join(', ') + '.',
                      type: 'choice'
                    });
                  }
                }
              }
            }
          }
          // --- User existence (single/multi) ---
          if ((/User/i).test(f.TypeAsString)) {
            let validUserMap = userValidValueByField[internal];
            if (validUserMap && !(raw === null || raw === undefined || this.cleanCell(raw) === '')) {
              let isMultiUser2 = (/Multi/i).test(f.TypeAsString);
              if (isMultiUser2) {
                let uparts = splitTokens(raw);
                for (let up = 0; up < uparts.length; up++) {
                  let uv = uparts[up];
                  let ulow = uv.toLowerCase();
                  if (validUserMap[ulow] !== true) {
                    // Custom message based on format
                    let msg = 'User "' + uv + '" not found.';
                    if (uv.toLowerCase().indexOf('i:0#.w|') !== 0) {
                      msg = 'Invalid format for "' + uv + '". Must be raw claim (e.g. "i:0#.w|domain\\user").';
                    }
                    issues.push({ row: dispRow, column: (f.Title || internal), detail: msg, type: 'user' });
                  }
                }
              } else {
                let uv1 = this.cleanCell(raw);
                if (uv1) {
                  let ul1 = uv1.toLowerCase();
                  if (validUserMap[ul1] !== true) {
                    issues.push({ row: dispRow, column: (f.Title || internal), detail: 'User "' + uv1 + '" not found.', type: 'user' });
                  }
                }
              }
            }
          }
          // --- Unique in selected range (sheet-level) ---
          if (uniqueTrack[internal]) {
            let strVal = String(raw || '').trim();
            if (strVal) {
              let seenMap = uniqueTrack[internal].seen;
              if (seenMap[strVal] >= 1) {
                issues.push({ row: dispRow, column: (f.Title || internal), detail: 'Duplicate value "' + strVal + '".', type: 'unique' });
              }
              seenMap[strVal] = (seenMap[strVal] || 0) + 1;
            }
          }
        }
      }
      // === SHAREPOINT-LEVEL UNIQUENESS CHECK (EnforceUniqueValues) ===
      // Only run if we have any unique fields with values
      const uniqueFieldsWithValues: { field: any; values: string[] }[] = [];
      for (const internal in uniqueTrack) {
        if (!Object.prototype.hasOwnProperty.call(uniqueTrack, internal)) continue;
        const track = uniqueTrack[internal];
        const values = Object.keys(track.seen).filter(v => track.seen[v] > 0);
        if (values.length > 0) {
          let field: IFieldInfo | undefined;
          for (let i = 0; i < fields.length; i++) {
            if (fields[i].InternalName === internal) {
              field = fields[i];
              break;
            }
          }
          if (field) {
            uniqueFieldsWithValues.push({ field, values });
          }
        }
      }
      if (uniqueFieldsWithValues.length > 0 && selectedList) {
        for (const { field, values } of uniqueFieldsWithValues) {
          const internal = field.InternalName;
          try {
            // Build OData filter: (Field eq 'val1') or (Field eq 'val2') ...
            const filterParts = values.map(v => `${internal} eq '${escapeOdataValue(v)}'`);
            const filter = filterParts.join(' or ');
            const existing = await this.sp.web.lists
              .getByTitle(selectedList)
              .items
              .select('Id', internal)
              .filter(filter)
              .top(5000)
              ();
            const existingValues = new Set(existing.map((e: any) => String(e[internal] || '')));
            // Re-scan rows to map back to Excel row number
            for (let idx = from; idx <= to; idx++) {
              const row = excelData[idx];
              const excelCol = excelMapByField[internal];
              if (!excelCol) continue;
              const rawVal = row[excelCol];
              const strVal = String(rawVal || '').trim();
              if (strVal && existingValues.has(strVal)) {
                const dispRow = (idx - from + 1);
                issues.push({
                  row: dispRow,
                  column: (field.Title || internal),
                  detail: `Value "${strVal}" already exists in SharePoint list (unique field).`,
                  type: 'unique'
                });
              }
            }
          } catch (err: any) {
            console.warn(`Failed to validate uniqueness in SP for field ${internal}:`, err);
            void LoggerService.log(selectedList || 'System', 'runDataValidation-unique - ' + this.state.jobName, 'Medium', this.state.mode, err.message);
          }
        }
      }
      // Summarize counts
      let counts = { required: 0, number: 0, choice: 0, unique: 0, user: 0, total: issues.length };
      for (let k = 0; k < issues.length; k++) {
        let t = issues[k].type;
        if (t === 'required') counts.required++;
        else if (t === 'number') counts.number++;
        else if (t === 'choice') counts.choice++;
        else if (t === 'unique') counts.unique++;
        else if (t === 'user') counts.user++;
      }
      this.setState({
        validationIssues: issues,
        validationRan: true,
        issueCounts: counts,
        validationError: issues.length ? '' : ''
      });
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'runDataValidation - ' + this.state.jobName, 'High', this.state.mode, e.message);
      this.setState({ validationError: 'Validation Failed: ' + (e as any).message });
    }
  }
  private cleanCell = (val: any): string => {
    try {
      if (val == null) return '';
      const str = String(val);
      return str.replace(/<[^>]*>/g, '').trim();
    } catch (e: any) {
      void LoggerService.log('System', 'cleanCell - ' + this.state.jobName, 'Low', this.state.mode, e.message);
      return '';
    }
  }
  private downloadValidationErrors = () => {
    try {
      const issues = this.state.validationIssues;
      if (!issues || !issues.length) return;
      const rows = issues.map(function (x) {
        return { Row: x.row, Column: x.column, Type: x.type, Detail: x.detail };
      });
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, 'Errors');
      const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([buf], { type: 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'validation_errors.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 0);
    } catch (e: any) {
      void LoggerService.log('System', 'downloadValidationErrors - ' + this.state.jobName, 'Medium', this.state.mode, e.message);
    }
  }
  // -------------- Build payload --------------
  // Build the SharePoint payload from one Excel row (NO validation here)
  private buildItemData = async (row: any): Promise<any> => {
    try {
      const fields = this.state.fields;
      const excelMapByField = this.state.excelMapByField;
      const selectedByField = this.state.selectedByField;
      const data: any = {};
      // local helper for tokenizing multi-value cells like "A;#B", "A;B", "A,B"
      const splitChoices = (raw: any): string[] => {
        if (raw === null || raw === undefined) return [];
        return String(raw)
          .replace(/;#/g, ';')
          .split(/[;,]/)
          .map(function (s) { return s.trim(); })
          .filter(function (s) { return !!s; });
      };
      for (let i = 0; i < fields.length; i++) {
        const fld = fields[i];
        const internal = fld.InternalName;
        // only include fields that are checked and mapped
        if (!selectedByField[internal]) continue;
        const excelCol = excelMapByField[internal];
        if (!excelCol) continue;
        let val: any = row[excelCol];
        // ----- DateTime -----
        if (fld.TypeAsString === 'DateTime' && val != null && val !== '') {
          if (val instanceof Date) {
            val = val.toISOString();
          } else if (typeof val === 'number') {
            // Excel serial dates; adjust for the 1900 leap bug using the ">=60" rule
            const offset = val >= 60 ? val - 1 : val;
            val = new Date(Date.UTC(1899, 11, 31) + offset * 86400000).toISOString();
          } else if (typeof val === 'string') {
            const parts = val.split(' ');
            const datePart = parts[0];
            const timePart = parts.length > 1 ? parts[1] : '';
            const dmy = datePart.split('/');
            if (dmy.length === 3) {
              const dd = parseInt(dmy[0], 10);
              const mm = parseInt(dmy[1], 10);
              const yyyy = parseInt(dmy[2], 10);
              const d = new Date(Date.UTC(yyyy, mm - 1, dd));
              if (timePart) {
                const hm = timePart.split(':');
                if (hm.length >= 2) {
                  const hh = parseInt(hm[0], 10);
                  const mi = parseInt(hm[1], 10);
                  if (!isNaN(hh) && !isNaN(mi)) d.setUTCHours(hh, mi, 0, 0);
                }
              }
              val = d.toISOString();
            }
          }
          data[internal] = val;
          continue;
        }
        // ----- Lookup / Multi Lookup -----
        if ((/Lookup/i).test(fld.TypeAsString) && fld.LookupList && val) {
          const listGuid = stripBraces(fld.LookupList);
          const displayFld = fld.LookupField || 'Title';
          const rawValues = fld.AllowMultipleValues
            ? String(val).split(/[;,]/).map(function (v) { return v.trim(); }).filter(function (v) { return !!v; })
            : [String(val).trim()];
          const ids: number[] = [];
          for (let r = 0; r < rawValues.length; r++) {
            const raw = rawValues[r];
            const items = await this.sp.web.lists
              .getById(listGuid)
              .items
              .select('Id')
              .filter(displayFld + " eq '" + escapeOdataValue(raw) + "'")
              .top(1)
              ();
            if (items.length) ids.push(items[0].Id);
          }
          data[internal + 'Id'] = fld.AllowMultipleValues ? { results: ids } : (ids.length ? ids[0] : null);
          continue;
        }
        // ----- Choice / MultiChoice (payload build; validator already checked) -----
        if ((/Choice/i).test(fld.TypeAsString)) {
          const isMulti = (/MultiChoice/i).test(fld.TypeAsString) || fld.AllowMultipleValues;
          // Try to use normalized allowed list if present (optional, safe fallback to [])
          const allowedList = (this.state.fields || [])
            .filter(ff => ff.InternalName === fld.InternalName)
            .map(ff => (ff as any).Choices as string[] || [])
            .shift() || [];
          const tokens = splitChoiceCell(val, !!isMulti, allowedList);
          if (isMulti) {
            data[internal] = { results: tokens.map(clean) };
          } else {
            data[internal] = tokens.length ? clean(tokens[0]) : '';
          }
          continue;
        }
        // ----- User / Multi User -----
        if ((/User/i).test(fld.TypeAsString) && val) {
          const isMulti = (/Multi/i).test(fld.TypeAsString);
          const rawValues = isMulti
            ? String(val).split(/[;,]/).map(function (v) { return v.trim(); }).filter(function (v) { return !!v; })
            : [String(val).trim()];
          const ids: number[] = [];
          for (let r = 0; r < rawValues.length; r++) {
            const raw = rawValues[r];
            let userId: number | null = null;
            try {
              const ensureRes: any = await this.sp.web.ensureUser(raw);
              userId = ensureRes.data.Id;
            } catch (err: any) {
              const found = await this.sp.web.siteUsers
                .filter("Title eq '" + escapeOdataValue(raw) + "'")
                .select('Id')
                .top(1)
                ();
              if (found.length) userId = found[0].Id;
            }
            if (userId !== null) ids.push(userId);
          }
          data[internal + 'Id'] = isMulti ? { results: ids } : (ids.length ? ids[0] : null);
          continue;
        }
        // ----- Hyperlink -----
        if ((/URL/i).test(fld.TypeAsString) && val) {
          const parts = String(val).split('|');
          const urlPart = parts[0] ? parts[0].trim() : '';
          const descPart = (parts.length > 1 && parts[1]) ? parts[1].trim() : urlPart;
          data[internal] = { Url: urlPart, Description: descPart };
          continue;
        }
        // ----- Fallback: text/number/etc. -----
        data[internal] = val;
      }
      return data;
    } catch (e: any) {
      void LoggerService.log(this.state.selectedList || 'System', 'buildItemData - ' + this.state.jobName, 'High', this.state.mode, e.message);
      throw e;
    }
  }
  // -------------- Helper: libraries & folders --------------
  // Ensures the library with the given title exists (creates if needed)
  // Does not check permissions; caller should do that
  private async ensureLibrary(title: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(title).select('Id')();
    } catch (e: any) {
      try {
        await this.sp.web.lists.add(title, title + ' library', 101, true);
      } catch (e2: any) {
        void LoggerService.log('System', 'ensureLibrary - ' + title, 'High', 'N/A', e2.message);
      }
    }
  }
  // Ensures the full folder path exists under the given server-relative root (library)
  // Returns the final folder object
  private async ensureFolderPath(serverRelRoot: string, path: string): Promise<any> {
    try {
      const parts = path.split('/').filter(function (p) { return !!p; });
      let currentUrl = serverRelRoot;
      for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        try {
          await this.sp.web.getFolderByServerRelativePath(currentUrl).folders.addUsingPath(part);
        } catch (e: any) {
          // Folder likely exists, log as low severity info if critical debugging needed
          void LoggerService.log('System', 'ensureFolderPath(loop) - ' + path, 'Low', 'N/A', e.message);
        }
        currentUrl = currentUrl + '/' + encodeURIComponent(part);
      }
      return this.sp.web.getFolderByServerRelativePath(currentUrl);
    } catch (e: any) {
      void LoggerService.log('System', 'ensureFolderPath(outer) - ' + path, 'Medium', 'N/A', e.message);
      throw e;
    }
  }

  // -------------- Import / Update (New or Resume) --------------
  private handleImport = async (): Promise<void> => {
    try {
      // 1. SHOW SPINNER IMMEDIATELY
      this.setState({ isLoading: true, validationError: '' });

      const {
        selectedList, excelData, mode, primaryField, fields,
        excelMapByField, selectedByField, startRow, endRow,
        jobMode, selectedJobId, sourceFileName, sourceFileBuffer
      } = this.state;

      // 2. BASIC VALIDATION (If fail, stop loading)
      if (!selectedList || excelData.length === 0) {
        this.setState({ isLoading: false });
        return;
      }

      // Required mapping check
      const missingRequired: string[] = [];
      for (let i = 0; i < fields.length; i++) {
        const f = fields[i];
        if (f.Required) {
          const mapped = !!excelMapByField[f.InternalName];
          const checked = !!selectedByField[f.InternalName];
          if (!mapped || !checked) missingRequired.push(f.Title || f.InternalName);
        }
      }

      if (missingRequired.length) {
        this.setState({
          isLoading: false, // Stop spinner
          validationError: 'Please map and check all required fields: ' + missingRequired.join(", ")
        });
        this.goToStep(2);
        return;
      }

      // Update mode key checks
      if (mode === 'update') {
        if (!primaryField) {
          this.setState({ isLoading: false, validationError: 'Select a primary key (Excel column).' });
          this.goToStep(3);
          return;
        }
        // Check for duplicates in the excel range itself
        const fromK = Math.max(1, Math.min(startRow, excelData.length)) - 1;
        const toK = Math.max(fromK, Math.min(endRow, excelData.length)) - 1;
        const seen: { [k: string]: boolean } = {};
        for (let i = fromK; i <= toK; i++) {
          const keyVal = String(excelData[i][primaryField] || '').trim();
          if (seen[keyVal]) {
            this.setState({ isLoading: false, validationError: 'Duplicate key "' + keyVal + '" in selected range.' });
            this.goToStep(3);
            return;
          }
          seen[keyVal] = true;
        }
      }

      // 3. SETUP JOB IDENTITY
      let jobName = (this.state.jobName || '').trim();
      let historyItemId: number | null = null;
      let baseSuccess = 0, baseFailure = 0, basePlanned = 0;

      if (jobMode === 'resume') {
        if (!selectedJobId) {
          this.setState({ isLoading: false, validationError: 'Please pick an existing job to resume.' });
          this.goToStep(1);
          return;
        }
        try {
          const item = await this.sp.web.lists.getByTitle(this.props.metricsListTitle).items
            .getById(selectedJobId)
            .select('Id,Title,SuccessCount,FailureCount,ItemstoImport')
            ();
          historyItemId = item.Id;
          jobName = item.Title || '';
          baseSuccess = (item.SuccessCount || 0);
          baseFailure = (item.FailureCount || 0);
          basePlanned = (item.ItemstoImport || 0);

          // Mark as running
          try { await this.sp.web.lists.getByTitle(this.props.metricsListTitle).items.getById(historyItemId!).update({ Status: 'Running' }); } catch (_) { }

          this.setState({
            jobName: jobName,
            historyItemId: historyItemId,
            baseSuccessCount: baseSuccess,
            baseFailureCount: baseFailure,
            basePlannedCount: basePlanned
          });
        } catch (e: any) {
          console.warn('Failed to load selected DataSycnHistory item', e);
          void LoggerService.log(this.props.metricsListTitle, 'handleImport-resumeLoad', 'High', mode, e.message);
          this.setState({ isLoading: false, validationError: 'Could not load the selected job. Try again.' });
          return;
        }
      } else {
        // NEW JOB
        if (!jobName) {
          this.setState({ isLoading: false, validationError: 'Job Name is required.' });
          this.goToStep(1);
          return;
        }
        // Unique Check
        try {
          const dup = await this.sp.web.lists.getByTitle(this.props.metricsListTitle).items
            .filter("Title eq '" + (jobName.replace(/'/g, "''")) + "'")
            .top(1)();
          if (dup && dup.length) {
            this.setState({ isLoading: false, validationError: 'Job Name "' + jobName + '" already exists.' });
            this.goToStep(1);
            return;
          }
        } catch (e: any) {
          // Ignore check errors
        }
      }

      // 4. PREPARE FOLDERS & FILES
      await this.ensureLibrary(this.props.metricsLibTitle);
      const libRoot = this.props.metricsLibTitle;
      await this.ensureFolderPath(libRoot, jobName);
      await this.ensureFolderPath(libRoot, jobName + '/Source');
      await this.ensureFolderPath(libRoot, jobName + '/Success');
      await this.ensureFolderPath(libRoot, jobName + '/Failure');

      // Upload Source
      let sourceUrl = '';
      if (sourceFileBuffer && sourceFileName) {
        try {
          const srcFolder = await this.ensureFolderPath(libRoot, jobName + '/Source');
          await srcFolder.files.add(sourceFileName, sourceFileBuffer, true);
          sourceUrl = window.location.origin + '/' + libRoot + '/' + encodeURIComponent(jobName) + '/Source/' + encodeURIComponent(sourceFileName);
          this.setState({ sourceFileUrl: sourceUrl });
        } catch (e: any) {
          void LoggerService.log(this.props.metricsLibTitle, 'handleImport-uploadSource', 'Medium', mode, e.message);
        }
      }

      // Create/Update History Item
      if (jobMode === 'new') {
        try {
          const start = new Date();
          this.jobStart = start;
          const addRes = await this.sp.web.lists.getByTitle(this.props.metricsListTitle).items.add({
            Title: jobName,
            DataSycnList: selectedList || '',
            OperationType: mode === 'add' ? 'Add' : 'Update',
            JobStartTime: start.toISOString(),
            ItemstoImport: 0,
            SuccessCount: 0,
            FailureCount: 0,
            Status: 'Running',
            SourceFile: sourceUrl ? { Url: sourceUrl, Description: sourceFileName || 'Source file' } : null
          });
          historyItemId = addRes.data.Id;
          this.setState({ historyItemId });
        } catch (e: any) {
          console.warn("Failed to create history item", e);
        }
      } else if (historyItemId && sourceUrl) {
        try {
          await this.sp.web.lists.getByTitle(this.props.metricsListTitle).items.getById(historyItemId).update({
            SourceFile: { Url: sourceUrl, Description: sourceFileName || 'Source file' }
          });
        } catch (e: any) { }
      }

      // Re-validate Data
      await this.runDataValidation();
      if (this.state.validationIssues.length > 0) {
        this.setState({ isLoading: false }); // Stop loading to show errors
        this.goToStep(4);
        return;
      }

      // 5. PERMISSION PROBE
      try {
        const probeList = this.sp.web.lists.getByTitle(selectedList);
        if (mode === 'add') {
          const probe = await probeList.items.add({ Title: '_perm_probe_' + Date.now() });
          try { await probeList.items.getById(probe.data.Id).recycle(); } catch (_) { }
        } else {
          // Edit check
          const any = await probeList.items.select('Id').top(1)();
          if (any && any.length) {
            await probeList.items.getById(any[0].Id).update({});
          }
        }
      } catch (permErr) {
        this.setState({ isLoading: false, validationError: 'Permission denied on list: ' + selectedList });
        if (this.props.showUserAlerts) void Swal.fire('Permission denied', 'You do not have rights to modify this list.', 'error');
        return;
      }

      // 6. START PROCESSING
      const from = Math.max(1, Math.min(startRow, excelData.length)) - 1;
      const to = Math.max(from, Math.min(endRow, excelData.length)) - 1;
      const plannedTotalThisRun = to - from + 1;

      // Update planned count
      if (historyItemId) {
        try {
          await this.sp.web.lists.getByTitle(this.props.metricsListTitle).items.getById(historyItemId).update({
            ItemstoImport: (this.state.basePlannedCount || 0) + plannedTotalThisRun
          });
        } catch (e: any) { }
      }

      this.setState({ progress: 0, completed: 0, total: plannedTotalThisRun });

      const list = this.sp.web.lists.getByTitle(selectedList);
      const successRows: any[] = [];
      const failureRows: any[] = [];

      // Find SP Field for Primary Key (Update Mode)
      let spFieldForExcel: string | null = null;
      if (mode === 'update') {
        for (const internal in excelMapByField) {
          if (excelMapByField[internal] === primaryField) {
            spFieldForExcel = internal;
            break;
          }
        }
      }

      // --- LOOP ---
      for (let idx = from; idx <= to; idx++) {
        const row = excelData[idx];

        try {
          const payload = await this.buildItemData(row); // Build Payload

          if (mode === 'update') {
            if (!spFieldForExcel) throw new Error("Primary key mapping missing");
            const rawKey = String(row[primaryField] || '').replace(/'/g, "''");
            const existing = await list.items.filter(spFieldForExcel + " eq '" + rawKey + "'").top(1)();

            if (existing.length > 0) {
              await list.items.getById(existing[0].Id).update(payload);
            } else {
              await list.items.add(payload);
            }
          } else {
            await list.items.add(payload);
          }
          successRows.push(this.extendRowWithStatus(row, 'Success'));
        } catch (importError) {
          const err = importError as any;
          const msg = err.message || JSON.stringify(err);
          const extRow = this.extendRowWithStatus(row, 'Failure', msg);
          extRow['Timestamp'] = new Date().toISOString();
          failureRows.push(extRow);
        }

        // Update Progress
        const completed = idx - from + 1;
        const percent = Math.round((completed / plannedTotalThisRun) * 100);
        this.setState({ progress: percent, completed: completed });
      }

      // 7. FINALIZE (Upload Results)
      const ts = new Date().toISOString().replace(/[-:]/g, '').split('.')[0];

      const makeWorkbook = (rows: any[]) => {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, 'Results');
        return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      };

      let successUrl = '', failureUrl = '';
      if (successRows.length) {
        const buf = makeWorkbook(successRows);
        const fname = `success_${this.state.selectedSheet}_${ts}.xlsx`;
        const folder = await this.ensureFolderPath(this.props.metricsLibTitle, jobName + '/Success');
        await folder.files.add(fname, buf, true);
        successUrl = window.location.origin + '/' + this.props.metricsLibTitle + '/' + encodeURIComponent(jobName) + '/Success/' + encodeURIComponent(fname);
      }
      if (failureRows.length) {
        const buf = makeWorkbook(failureRows);
        const fname = `failure_${this.state.selectedSheet}_${ts}.xlsx`;
        const folder = await this.ensureFolderPath(this.props.metricsLibTitle, jobName + '/Failure');
        await folder.files.add(fname, buf, true);
        failureUrl = window.location.origin + '/' + this.props.metricsLibTitle + '/' + encodeURIComponent(jobName) + '/Failure/' + encodeURIComponent(fname);
      }

      // Update History Item Final Status
      if (historyItemId) {
        const totalSuccess = (this.state.baseSuccessCount || 0) + successRows.length;
        const totalFailure = (this.state.baseFailureCount || 0) + failureRows.length;

        let finalStatus = 'Completed';
        if (totalFailure > 0 && totalSuccess === 0) finalStatus = 'Failed';
        else if (totalFailure > 0) finalStatus = 'Completed with errors';

        await this.sp.web.lists.getByTitle(this.props.metricsListTitle).items.getById(historyItemId).update({
          JobEndTime: new Date().toISOString(),
          SuccessCount: totalSuccess,
          FailureCount: totalFailure,
          Status: finalStatus,
          SuccessFile: successUrl ? { Url: successUrl, Description: 'Success File' } : null,
          FailureFile: failureUrl ? { Url: failureUrl, Description: 'Failure File' } : null
        });
      }

      // 8. UPDATE UI & SHOW ALERT
      this.setState({
        isLoading: false,
        progress: 100, // Force 100% to show buttons
        successCount: successRows.length,
        failureCount: failureRows.length,
        successUrl,
        failureUrl
      });

      // Show Simple Notification (Only if alerts enabled)
      if (this.props.showUserAlerts) {
        await void Swal.fire({
          title: 'Import Finished',
          html: `
               <p>${successRows.length} succeeded, ${failureRows.length} failed.</p>
               ${failureRows.length > 0 ? '<p style="color:red">Please check the failure log.</p>' : ''}
            `,
          icon: failureRows.length > 0 ? 'warning' : 'success',
          confirmButtonText: 'OK' // Just closes the popup
        });
      }

    } catch (e: any) {
      this.setState({ isLoading: false, validationError: 'Critical Error: ' + e.message });
      void LoggerService.log('System', 'handleImport', 'High', this.state.mode, e.message);
    }
  }

  private extractSheetFromFailureFileName = (fileName: string): string | null => {
    try {
      const match = fileName.match(/^failure_([^_]+)_/);
      return match ? match[1] : null;
    } catch (e: any) {
      void LoggerService.log('System', 'extractSheetFromFailureFileName', 'Low', 'N/A', e.message);
      return null;
    }
  }
  private async moveSourceFileToSuccess(jobName: string, sourceFileName: string): Promise<void> {
    try {
      const libRoot = this.props.metricsLibTitle;
      const failureFolder = await this.ensureFolderPath(libRoot, `${jobName}/Failure`);
      const successFolder = await this.ensureFolderPath(libRoot, `${jobName}/Success`);
      const filePath = failureFolder.ServerRelativeUrl + '/' + sourceFileName;
      const targetPath = successFolder.ServerRelativeUrl + '/' + sourceFileName;
      const file = this.sp.web.getFileByServerRelativePath(filePath);
      await (file as any).moveToPath(targetPath, { Overwrite: true });
    } catch (err: any) {
      console.warn('Failed to move source file to Success folder', err);
      void LoggerService.log(this.props.metricsLibTitle, 'moveSourceFileToSuccess - ' + jobName, 'Medium', this.state.mode, err.message);
    }
  }
  //  -------------- Exit to Dashboard --------------
  private handleExitToDashboard = () => {
    try {
      if (this.props.onExitToDashboard) {
        this.props.onExitToDashboard();
      } else {
        this.setState({ currentStep: 1 });
      }
    } catch (e: any) {
      void LoggerService.log('System', 'handleExitToDashboard', 'Medium', 'N/A', e.message);
    }
  }
  // Extend original Excel row with status and optional error details
  private extendRowWithStatus = (row: any, dataSyncMigrationstatus: string, dataSyncErrorDetails?: string) => {
    try {
      const copy: any = {};
      for (const k in row) if (Object.prototype.hasOwnProperty.call(row, k)) copy[k] = row[k];
      copy.dataSyncMigrationstatus = dataSyncMigrationstatus;
      if (dataSyncErrorDetails) copy.dataSyncErrorDetails = dataSyncErrorDetails;
      return copy;
    } catch (e: any) {
      void LoggerService.log('System', 'extendRowWithStatus', 'Low', 'N/A', e.message);
      return row;
    }
  }

  public render() {
    try {
      const {
        lists, sheetOptions, selectedSheet, columns, fields,
        selectedList, fileError, mode, primaryField, validationError,
        isLoading, progress, completed, total,
        excelMapByField, selectedByField, startRow, endRow, currentStep,
        validationIssues, validationRan,
        jobMode, existingJobs, selectedJobId, jobName
      } = this.state;

      // 1. Calculate columns used globally (outer scope)
      const usedExcelCols: string[] = [];
      for (const k in excelMapByField) {
        if (Object.prototype.hasOwnProperty.call(excelMapByField, k) && excelMapByField[k]) {
          usedExcelCols.push(excelMapByField[k]);
        }
      }

      const steps = [
        { n: 1, title: 'Upload & Sheet' },
        { n: 2, title: 'List & Mapping' },
        { n: 3, title: 'Operation Type' },
        { n: 4, title: 'Validate & Preview' },
        { n: 5, title: 'Review & Run' }
      ];

      // Preview logic
      const previewFields: IFieldInfo[] = [];
      for (let i = 0; i < fields.length; i++) {
        const f = fields[i];
        if (selectedByField[f.InternalName] && excelMapByField[f.InternalName]) {
          previewFields.push(f);
        }
      }
      const fromIdx = Math.max(1, Math.min(startRow, this.state.excelData.length)) - 1;
      const toIdx = Math.max(fromIdx, Math.min(endRow, this.state.excelData.length)) - 1;
      const maxPreview = 15;
      const previewEnd = Math.min(toIdx, fromIdx + maxPreview - 1);

      const css: any = styles; // Bypass strict SCSS types

      return (
        <div className={css.powerDataSyncContainer}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
            <h3 style={{ margin: 0 }}>Power Data Synchronizer</h3>
          </div>

          {/* Stepper header */}
          <div className={css.wizard} style={{ justifyContent: 'space-between' }}>
            {steps.map((s, i) => {
              const isActive = currentStep === s.n;
              const isDone = currentStep > s.n;
              const cls = (css.wizardStep || '') + ' ' + (isActive ? (css.wizardStepActive || '') : '') + ' ' + (isDone ? (css.wizardStepDone || '') : '');

              const handleStepClick = () => {
                if (s.n <= currentStep) {
                  this.goToStep(s.n);
                  return;
                }
                if (this.validateStep(currentStep)) {
                  if (s.n === 4 && currentStep <= 3) { void this.runDataValidation(); }
                  this.goToStep(s.n);
                }
              };

              return (
                <div key={'step-' + s.n} className={cls} onClick={handleStepClick} title={s.title}>
                  <span className={css.wizardIndex}>{s.n}</span>
                  <span className={css.wizardTitle}>{s.title}</span>
                  {i < steps.length - 1 ? <span className={css.wizardSeparator} /> : null}
                </div>
              );
            })}

            <div style={{ cursor: 'pointer', marginLeft: 'auto' }} onClick={this.handleExitToDashboard} title="Back to Dashboard">
              <i className="ms-Icon ms-Icon--Home" style={{ fontSize: 20 }} />
            </div>
          </div>

          {fileError ? <div className={styles.errorText}>{fileError}</div> : null}

          {/* ===== STEP 1 ===== */}
          {currentStep === 0 ? (
            <div>
              <div className={styles.formGroup}>
                <label>Upload Excel / CSV</label>
                <input type="file" accept=".xlsx,.xls,.csv" onChange={this.handleFileUpload} />
              </div>
              {sheetOptions.length > 0 ? (
                <div className={styles.formGroup}>
                  <label>Select a Sheet</label>
                  <select value={selectedSheet} onChange={this.handleSheetChange}>
                    {sheetOptions.map(function (s) { return <option key={s} value={s}>{s}</option>; })}
                  </select>
                </div>
              ) : null}
            </div>
          ) : null}

          {/* ===== STEP 1 (Main) ===== */}
          {currentStep === 1 ? (
            <div>
              <div className={styles.formGroup}>
                <label>Job Mode</label>
                <div>
                  <label style={{ marginRight: 16 }}>
                    <input type="radio" name="jobMode" checked={jobMode === 'new'} onChange={() => this.handleJobModeChange('new')} /> New
                  </label>
                  <label>
                    <input type="radio" name="jobMode" checked={jobMode === 'resume'} onChange={() => this.handleJobModeChange('resume')} /> Resume
                  </label>
                </div>
              </div>

              {jobMode === 'new' ? (
                <div className={styles.formGroup}>
                  <label>Job Name <span style={{ color: '#d32f2f' }}>*</span></label>
                  <div style={{ position: 'relative' }}>
                    <input
                      type="text"
                      value={jobName}
                      onChange={e => {
                        const val = e.target.value;
                        this.setState({ jobName: val });
                        this.checkJobName(val);
                      }}
                      placeholder={`e.g., ${new Date().toLocaleString('default', { month: 'long' })} ${new Date().getFullYear()} Batch 1`}
                      style={{ width: '100%', padding: '8px 10px', border: this.state.jobNameError ? '1px solid #d32f2f' : '1px solid #ccc', borderRadius: 4 }}
                    />
                    {this.state.jobNameChecking && (
                      <div style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', fontSize: 14, color: '#666' }}>
                        checking...
                      </div>
                    )}
                  </div>
                  {this.state.jobNameError && (
                    <div style={{ color: '#d32f2f', fontSize: 13, marginTop: 4 }}>{this.state.jobNameError}</div>
                  )}
                </div>
              ) : (
                <div className={styles.formGroup}>
                  <label>Pick a failed job to resume</label>
                  <select value={selectedJobId || ''} onChange={e => this.handleSelectExistingJob(e.target.value)} onFocus={() => this.loadExistingJobs()}>
                    <option value="">-- Select Failed Job --</option>
                    {existingJobs.map(j => (
                      <option key={j.Id} value={String(j.Id)}>{j.Title} ({j.Status})</option>
                    ))}
                  </select>
                </div>
              )}

              <div className={styles.formGroup}>
                <label>Upload Excel / CSV <span className={styles.muted}>(required)</span></label>
                <input type="file" accept=".xlsx,.xls,.csv" onChange={this.handleFileUpload} />
              </div>

              {sheetOptions.length > 0 ? (
                <div className={styles.formGroup}>
                  <label>Select a Sheet</label>
                  <select value={selectedSheet} onChange={this.handleSheetChange}>
                    {sheetOptions.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
              ) : null}
            </div>
          ) : null}

          {/* ===== STEP 2 ===== */}
          {currentStep === 2 ? (
            <div>
              <div className={styles.formGroup}>
                <label>Select a SharePoint List</label>
                <select value={selectedList} onChange={this.handleListChange}>
                  <option value="">--Select List--</option>
                  {lists.map(l => <option key={l} value={l}>{l}</option>)}
                </select>
              </div>

              {(selectedSheet && selectedList && columns.length > 0 && fields.length > 0) ? (
                <table className={styles.mappingTable}>
                  <thead>
                    <tr>
                      <th>SharePoint Field</th>
                      <th>Excel Column</th>
                    </tr>
                  </thead>
                  <tbody>
                    {fields.map(f => {
                      const isRequired = f.Required === true;
                      const isUnique = f.EnforceUniqueValues === true;
                      const mappedCol = excelMapByField[f.InternalName] || '';
                      const isChecked = !!selectedByField[f.InternalName];

                      //  Variable Shadowing - renamed to 'otherMappedCols'
                      const otherMappedCols: string[] = [];
                      for (const key in excelMapByField) {
                        if (Object.prototype.hasOwnProperty.call(excelMapByField, key) && key !== f.InternalName && excelMapByField[key]) {
                          otherMappedCols.push(excelMapByField[key]);
                        }
                      }

                      return (
                        <tr key={f.InternalName} className={isRequired ? styles.requiredRow : undefined}>
                          <td className={styles.colNameCell}>
                            <input type="checkbox" checked={isChecked} onChange={e => this.handleFieldCheck(f.InternalName, e.target.checked)} />
                            <span>
                              {(f.Title || f.InternalName)}
                              {isRequired && <span style={{ color: '#d32f2f', marginLeft: 4 }}>*</span>}
                              {isUnique && <span className={styles.uniqueBadge}>UNIQUE</span>}
                            </span>
                          </td>
                          <td className={styles.fieldSelectCell}>
                            <select value={mappedCol} onChange={e => this.handleExcelMapChange(f.InternalName, e.target.value)}>
                              <option value="">--Select--</option>
                              {columns.map(c => {
                                // Use the renamed variable here
                                const disabled = otherMappedCols.indexOf(c) !== -1 && c !== mappedCol;
                                return <option key={c} value={c} disabled={disabled}>{c}</option>;
                              })}
                            </select>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              ) : null}
            </div>
          ) : null}

          {/* ===== STEP 3 ===== */}
          {currentStep === 3 ? (
            <div>
              <div className={styles.formGroup}>
                <label>Operation</label><br />
                <label><input type="radio" checked={mode === 'add'} onChange={() => this.setState({ mode: 'add', validationError: '', primaryField: '' })} /> Add only</label>
                <label><input type="radio" checked={mode === 'update'} onChange={() => this.setState({ mode: 'update', validationError: '' })} /> Update</label>
              </div>
              {mode === 'update' ? (
                <div className={styles.formGroup}>
                  <label>Primary key (Excel column)</label>
                  <select value={primaryField} onChange={(e) => this.setState({ primaryField: e.target.value, validationError: '' })}>
                    <option value="">--Select--</option>
                    {columns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              ) : null}
              <div className={styles.formGroup}>
                <label>Import rows from:</label>
                <input type="number" min={1} max={endRow} value={startRow} onChange={(e) => this.setState({ startRow: parseInt(e.target.value, 10) })} />
                <span> to </span>
                <input type="number" min={1} max={endRow} value={endRow} onChange={(e) => this.setState({ endRow: parseInt(e.target.value, 10) })} />
                <span> (max {endRow} rows)</span>
              </div>
            </div>
          ) : null}

          {/* ===== STEP 4 ===== */}
          {currentStep === 4 ? (
            <div>
              <div className={styles.actionsRow}>
                <button className={styles.importButton} onClick={async () => { await this.runDataValidation(); }} disabled={isLoading} title="Re-run data validation">Re-Validate</button>
                <button className={styles.linkButton} onClick={this.downloadValidationErrors} disabled={!validationRan || !validationIssues.length} title="Download all validation errors as Excel">Download Errors (.xlsx)</button>
              </div>
              {validationRan ? (
                <div className={styles.validationSummary}>
                  <span className={styles.pill}>Total: {this.state.issueCounts.total}</span>
                  <span className={styles.pill}>Required: {this.state.issueCounts.required}</span>
                  <span className={styles.pill}>Number: {this.state.issueCounts.number}</span>
                  <span className={styles.pill}>Choice: {this.state.issueCounts.choice}</span>
                  <span className={styles.pill}>Unique: {this.state.issueCounts.unique}</span>
                  <span className={styles.pill}>User: {this.state.issueCounts.user}</span>
                </div>
              ) : null}
              {validationRan && validationIssues.length > 0 ? (
                <div className={styles.validationList}>
                  <ul>
                    {validationIssues.slice(0, 200).map((iss, idx) => (
                      <li key={'iss-' + idx}>Row {iss.row}: <em>{iss.column}</em> — {iss.detail}</li>
                    ))}
                  </ul>
                </div>
              ) : null}
              <div className={styles.formGroup}>
                <strong>Preview (first {Math.min(15, (previewEnd - fromIdx + 1) || 0)} rows):</strong>
                {(previewFields.length > 0 && previewEnd >= fromIdx) ? (
                  <div className={styles.previewTableWrapper}>
                    <table className={styles.previewTable}>
                      <thead>
                        <tr>{previewFields.map(f => <th key={'h-' + f.InternalName}>{f.Title || f.InternalName}</th>)}</tr>
                      </thead>
                      <tbody>
                        {this.state.excelData.slice(fromIdx, previewEnd + 1).map((row, rIdx) => (
                          <tr key={'r-' + rIdx}>
                            {previewFields.map(f => {
                              const col = excelMapByField[f.InternalName];
                              return <td key={'c-' + f.InternalName}>{String(col ? row[col] : '')}</td>;
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ) : <div>No fields selected or no rows in range.</div>}
              </div>
              {validationRan && validationIssues.length === 0 ? <div className={styles.successText}>No validation issues found. You can proceed.</div> : null}
            </div>
          ) : null}

          {/* ===== STEP 5 (Review & Run) ===== */}
          {currentStep === 5 ? (
            <div>
              <div className={styles.formGroup}>
                <strong>Review</strong>
                <ul>
                  <li>Sheet: <em>{selectedSheet || '—'}</em></li>
                  <li>List: <em>{selectedList || '—'}</em></li>
                  <li>Mode: <em>{mode === 'add' ? 'Add only' : 'Update'}</em> {mode === 'update' ? <span> (Key: <em>{primaryField || '—'}</em>)</span> : null}</li>
                  <li>Rows: <em>{startRow} to {endRow}</em></li>
                  {jobMode === 'resume' && selectedJobId ? <li>Job: <em>{jobName}</em> (resuming)</li> : null}
                </ul>
              </div>

              {isLoading ? (
                <div className={styles.progressBarContainer}>
                  <div className={styles.progressBar}>
                    <div className={styles.progressFill} style={{ width: String(progress) + '%' }} />
                  </div>
                  <div className={styles.progressText}>Importing {progress}% ({completed}/{total})</div>
                </div>
              ) : null}

              {/* ===== ACTIONS CONTAINER ===== */}
              {/*  Wrapped conditional buttons in a div to prevent TS2657/TS1109 errors */}
              <div style={{ marginTop: 20 }}>

                {/* 1. RUN BUTTON (Show only if NOT 100% complete) */}
                {(!isLoading && this.state.progress !== 100) && (
                  <button
                    className={styles.importButton}
                    onClick={this.handleImport}
                    disabled={
                      isLoading ||
                      !this.state.selectedList ||
                      this.state.columns.length === 0 ||
                      (this.state.jobMode === 'resume' && !this.state.selectedJobId) ||
                      (this.state.jobMode === 'new' && !this.state.jobName.trim()) ||
                      (this.state.jobMode === 'new' && !!this.state.jobNameError) ||
                      (this.state.jobMode === 'new' && this.state.jobNameChecking)
                    }
                  >
                    {isLoading ? 'Running...' : 'Run Import'}
                  </button>
                )}

                {/* 2. NAVIGATION BUTTONS (Show ONLY if 100% complete) */}
                {(!isLoading && this.state.progress === 100) && (
                  <div style={{ display: 'flex', gap: '10px' }}>
                    <button
                      className={styles.importButton}
                      style={{ backgroundColor: '#107c10', borderColor: '#107c10' }}
                      onClick={() => { if (this.props.onExitToDashboard) this.props.onExitToDashboard(); }}
                    >
                      Return to Dashboard
                    </button>
                    <button
                      className={styles.importButton}
                      style={{ backgroundColor: '#0078d4', borderColor: '#0078d4' }}
                      onClick={() => { if (this.props.onRunAnother) this.props.onRunAnother(); }}
                    >
                      Run Another Job
                    </button>
                  </div>
                )}
              </div>

            </div>
          ) : null}

          {/* Validation Error Message */}
          {validationError ? <div className={styles.errorText} style={{ marginTop: 8 }}>{validationError}</div> : null}

          {/* Loader Overlay */}
          {isLoading && (
            <div style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', backgroundColor: 'rgba(255,255,255,0.6)', zIndex: 9999, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'wait' }}>
              <div style={{ padding: 20, background: '#fff', borderRadius: 8, boxShadow: '0 2px 10px rgba(0,0,0,0.2)' }}>Processing...</div>
            </div>
          )}

          {/* Wizard Navigation (Bottom) */}
          <div className={styles.formGroup} style={{ display: 'flex', gap: 8, marginTop: 12 }}>
            <button className={styles.importButton} onClick={this.handleBack} disabled={currentStep === 1 || isLoading}>Back</button>
            {currentStep < 5 ? (
              <button
                className={styles.importButton}
                onClick={this.handleNext}
                disabled={
                  isLoading ||
                  (currentStep === 1 && (!this.state.jobName.trim() || !!this.state.jobNameError || this.state.jobNameChecking || !this.state.selectedSheet || this.state.columns.length === 0)) ||
                  (currentStep === 2 && (!this.state.selectedList || this.state.fields.length === 0)) ||
                  (currentStep === 4 && (!this.state.validationRan || this.state.validationIssues.length > 0))
                }
              >
                Next
              </button>
            ) : null}
          </div>
        </div >
      );
    } catch (e: any) {
      void LoggerService.log('System', 'render', 'High', 'UI', e.message);
      return <div>Critical Error in Render: {e.message}</div>;
    }
  }
}
