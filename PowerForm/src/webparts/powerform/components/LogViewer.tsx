import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { 
  DetailsList, 
  IColumn, 
  SelectionMode, 
  Selection, 
  CheckboxVisibility, 
  DetailsListLayoutMode 
} from '@fluentui/react/lib/DetailsList';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as XLSX from 'xlsx';
import { LoggerService, ILogItem } from './LoggerService';

export interface ILogViewerProps {
  isOpen: boolean;
  context: WebPartContext;
  currentListTitle: string;
  onDismiss: () => void;
}

export interface ILogViewerState {
  items: ILogItem[];
  allItems: ILogItem[];
  loading: boolean;
  selectionCount: number;
}

export class LogViewer extends React.Component<ILogViewerProps, ILogViewerState> {
  private _selection: Selection;

  constructor(props: ILogViewerProps) {
    super(props);
    
    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({ selectionCount: this._selection.getSelectedCount() });
      }
    });

    this.state = {
      items: [],
      allItems: [],
      loading: true,
      selectionCount: 0
    };
  }

  public componentDidUpdate(prevProps: ILogViewerProps): void {
    if (this.props.isOpen && !prevProps.isOpen) {
      void this._loadLogs();
    }
  }

  private async _loadLogs(): Promise<void> {
    try {
      this.setState({ loading: true });
      const logs = await LoggerService.getLogs(this.props.currentListTitle);
      this.setState({ items: logs, allItems: logs, loading: false });
    } catch (error:any) {
      void LoggerService.log('LogViewer-loadLogs', 'High', 'UI', error.message);
      this.setState({ loading: false });
    }
  }

  private _onSearch = (text: string): void => {
    if (!text) {
      this.setState({ items: this.state.allItems });
      return;
    }

    const lower = text.toLowerCase();
    const filtered = this.state.allItems.filter(i =>
      (i.Title?.toLowerCase().indexOf(lower) ?? -1) > -1 ||
      (i.Error?.toLowerCase().indexOf(lower) ?? -1) > -1 ||
      (i.Module?.toLowerCase().indexOf(lower) ?? -1) > -1 ||
      (i.Severity?.toLowerCase().indexOf(lower) ?? -1) > -1 ||
      (i.ErrorId?.toLowerCase().indexOf(lower) ?? -1) > -1
    );
    this.setState({ items: filtered });
  }

  private _exportLogs = (): void => {
    try {
      const selected = this._selection.getSelection() as ILogItem[];
      const dataToExport = selected.length > 0 ? selected : this.state.items;
      
      const cleanData = dataToExport.map(i => ({
        Date: i.Created ? new Date(i.Created).toLocaleString() : '',
        User: i.Author ? i.Author.Title : 'System',
        Module: i.Module,
        Page: i.Page,
        ItemId: i.ItemId,
        Severity: i.Severity,
        Error: i.Error,
        ReferenceID: i.ErrorId
      }));

      const ws = XLSX.utils.json_to_sheet(cleanData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "System Logs");
      XLSX.writeFile(wb, `PowerForm_Logs_${this.props.currentListTitle}.xlsx`);
    } catch (error:any) {
      void LoggerService.log('LogViewer-exportLogs', 'High', 'UI', error.message);
    }
  }

  public render(): React.ReactElement<ILogViewerProps> {
    const { items, loading, selectionCount } = this.state;

    const columns: IColumn[] = [
      { 
        key: 'col0', 
        name: 'Date', 
        fieldName: 'Created', 
        minWidth: 130, 
        maxWidth: 160, 
        onRender: (item: ILogItem) => item.Created ? new Date(item.Created).toLocaleString() : '-'
      },
      { key: 'col1', name: 'Module', fieldName: 'Module', minWidth: 70, maxWidth: 100 },
      {
        key: 'col2', name: 'Severity', fieldName: 'Severity', minWidth: 60, maxWidth: 80, 
        onRender: (item: ILogItem) => {
          const color = item.Severity === 'High' ? '#d13438' : item.Severity === 'Medium' ? '#ff8c00' : '#107c10';
          return <span style={{ color, fontWeight: 600 }}>{item.Severity}</span>;
        }
      },
      {
        key: 'colErrorId',
        name: 'Error ID',
        fieldName: 'ErrorId',
        minWidth: 120,
        maxWidth: 160,
        isResizable: true,
        onRender: (item: ILogItem) => <span style={{ background: '#f3f2f1', padding: '2px 4px', fontSize: '11px', fontFamily: 'monospace' }}>{item.ErrorId || '-'}</span>
      },
      { key: 'col3', name: 'Page', fieldName: 'Page', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'col4', name: 'Error Message', fieldName: 'Error', minWidth: 200, maxWidth: 500, isMultiline: true }
    ];

    return (
      <Panel
        isOpen={this.props.isOpen}
        onDismiss={this.props.onDismiss}
        type={PanelType.large}
        headerText="System Error Logs"
        closeButtonAriaLabel="Close"
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 15, gap: 10, alignItems: 'center' }}>
          <SearchBox
            placeholder="Filter logs by Module, Error, or ID..."
            onChange={(_, val) => this._onSearch(val || '')}
            style={{ flex: 1 }} 
          />
          <button
            onClick={this._exportLogs}
            style={{ 
              padding: '8px 20px', 
              background: '#0078d4', 
              color: 'white', 
              border: 'none', 
              cursor: 'pointer', 
              fontWeight: 600,
              borderRadius: '2px'
            }}
          >
            {selectionCount > 0 ? `Export Selected (${selectionCount})` : 'Export All'}
          </button>
        </div>

        {loading ? (
          <Spinner size={SpinnerSize.large} label="Fetching logs from SharePoint..." />
        ) : (
          <div style={{ height: '70vh', border: '1px solid #edebe9' }}>
            <DetailsList
              items={items}
              columns={columns}
              selection={this._selection}
              selectionMode={SelectionMode.multiple}
              checkboxVisibility={CheckboxVisibility.always}
              layoutMode={DetailsListLayoutMode.fixedColumns}
            />
            {items.length === 0 && <div style={{ padding: 40, textAlign: 'center', color: '#605e5c' }}>No log entries found for this list.</div>}
          </div>
        )}
      </Panel>
    );
  }
}