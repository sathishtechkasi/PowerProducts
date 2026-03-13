import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPowerDataSyncProps {
  // Standard SPFx Props
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
  
  // Custom Web Part Props
  context: WebPartContext;
  siteUrl: string;
  version?: string;
  onExitToDashboard?: () => void;
  onRunAnother?: () => void;
  resumeJobId?: number;
  
  // Dynamic properties from Property Pane
  metricsListTitle: string;
  metricsLibTitle: string;
  showUserAlerts: boolean;
  showHiddenLists: boolean;
}

export interface PowerDashboardProps {
  siteUrl: string;
  context: WebPartContext;
  metricsListTitle: string; 
  onNewJob: () => void;
  onResumeJob?: (id: number) => void;
}

export interface IValidationIssue {
  row: number;
  column: string;
  detail: string;
  type: 'required' | 'number' | 'choice' | 'unique' | 'user';
}

export interface IFieldInfo {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  LookupList?: string;
  LookupField?: string;
  AllowMultipleValues?: boolean;
  Required?: boolean;
  EnforceUniqueValues?: boolean;
  Choices?: string[];
  FillInChoice?: boolean;
}

export interface IExistingJob {
  Id: number;
  Title: string;
  Status: string;
}