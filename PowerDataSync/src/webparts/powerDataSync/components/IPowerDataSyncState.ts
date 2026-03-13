import { IFieldInfo, IValidationIssue, IExistingJob } from './IPowerDataSyncProps';
import { JobRow } from '../services/DataSyncHistoryService';
import * as XLSX from 'xlsx';

export interface IPowerDataSyncState {
  // Lists
  lists: string[];
  selectedList: string;
  // Wizard
  currentStep: number; 
  // Excel
  excelData: any[];
  columns: string[];
  sheetOptions: string[];
  selectedSheet: string;
  fileError: string;
  // SP fields
  fields: IFieldInfo[];
  // Mapping & selection
  excelMapByField: { [internal: string]: string };
  excelMapIndexByField: { [key: string]: number };
  selectedByField: { [internal: string]: boolean };
  // Mode
  mode: 'add' | 'update';
  primaryField: string;
  validationError: string;
  // Progress
  progress: number;
  completed: number;
  total: number;
  isLoading: boolean;
  workbook: XLSX.WorkBook | null;   
  // Results
  successUrl: string;
  failureUrl: string;
  successCount: number;
  failureCount: number;
  // Row range
  startRow: number;
  endRow: number;
  // Data validation
  validationIssues: IValidationIssue[];
  validationRan: boolean;
  issueCounts: {
    required: number;
    number: number;
    choice: number;
    unique: number;
    user: number;
    total: number;
  };
  showAllIssues: boolean;
  // Permissions on selected list
  canAdd: boolean;
  canEdit: boolean;
  // Job & source file
  jobName: string;
  jobMode: 'new' | 'resume';
  existingJobs: IExistingJob[];
  selectedJobId: number | null;
  jobNameError: string;      
  jobNameChecking: boolean;  
  sourceFileName: string;
  sourceFileBuffer: Uint8Array | null;
  sourceFileUrl: string;
  // History item
  historyItemId: number | null;
  // Resume: base counts to accumulate
  baseSuccessCount: number;
  baseFailureCount: number;
  basePlannedCount: number;
}

export interface PowerDashboardState {
  jobs: JobRow[];
  loading: boolean;
  // filters
  fTitle: string;
  fOwner: string;
  fStatus: string;
  fMinItems?: number;
  fMinSuccess?: number;
  fMinErrors?: number;
  // KPIs
  totalJobsCompleted: number;
  totalItemsSync: number;
  successRatio: number;
  totalJobsInitiated: number;
  totalJobsSuccess: number;
  totalJobsFailed: number;
  totalItemsSynchronized: number;
  totalItemsFailed: number;
  overallSuccessRatio: number;
  // Pagination
  page: number;
  pageSize: number;
  fListName: string;
}