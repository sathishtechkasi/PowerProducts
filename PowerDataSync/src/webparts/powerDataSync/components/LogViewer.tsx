import * as React from 'react';
import { 
    Panel, 
    PanelType, 
    DetailsList, 
    IColumn, 
    SelectionMode, 
    Selection, 
    CheckboxVisibility, 
    SearchBox, 
    Spinner, 
    SpinnerSize, 
    DetailsListLayoutMode 
} from '@fluentui/react';
import * as XLSX from 'xlsx';
import { LoggerService, ILogItem } from './LoggerService';

export interface ILogViewerProps {
    isOpen: boolean;
    context: any;
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
    private selection!: Selection;

    constructor(props: ILogViewerProps) {
        super(props);
        try {
            this.selection = new Selection({
                onSelectionChanged: () => {
                    try {
                        this.setState({ selectionCount: this.selection.getSelectedCount() });
                    } catch (error: any) {
                        LoggerService.log('Power DataSyncWeb Part', 'LogViewer-onSelectionChanged', 'Low', 'Config', error.message || JSON.stringify(error));
                    }
                }
            });
            this.state = {
                items: [],
                allItems: [],
                loading: true,
                selectionCount: 0
            };
        } catch (error: any) {
            LoggerService.log('Power DataSyncWeb Part', 'LogViewer-constructor', 'High', 'Config', error.message || JSON.stringify(error));
        }
    }

    public componentDidUpdate(prevProps: ILogViewerProps) {
        if (this.props.isOpen && !prevProps.isOpen) {
            this.loadLogs();
        }
    }

    private async loadLogs() {
        try {
            this.setState({ loading: true });
            const logs = await LoggerService.getLogs(this.props.context, this.props.currentListTitle);
            this.setState({ items: logs, allItems: logs, loading: false });
        } catch (error: any) {
            LoggerService.log('Power DataSyncWeb Part', 'LogViewer-loadLogs', 'High', 'Config', error.message || JSON.stringify(error));
            this.setState({ loading: false });
        }
    }

    private onSearch = (text: string) => {
        try {
            if (!text) {
                this.setState({ items: this.state.allItems });
                return;
            }
            const lower = text.toLowerCase();
            const filtered = this.state.allItems.filter(i =>
                (i.Title && i.Title.toLowerCase().indexOf(lower) > -1) ||
                (i.Error && i.Error.toLowerCase().indexOf(lower) > -1) ||
                (i.Module && i.Module.toLowerCase().indexOf(lower) > -1) ||
                (i.Severity && i.Severity.toLowerCase().indexOf(lower) > -1) ||
                (i.ErrorId && i.ErrorId.toLowerCase().indexOf(lower) > -1)
            );
            this.setState({ items: filtered });
        } catch (error: any) {
            LoggerService.log('Power DataSyncWeb Part', 'LogViewer-onSearch', 'Medium', 'Config', error.message || JSON.stringify(error));
        }
    }

    private exportLogs = () => {
        try {
            const selected = this.selection.getSelection() as ILogItem[];
            const dataToExport = selected.length > 0 ? selected : this.state.items;
            const cleanData = dataToExport.map(i => ({
                Date: i.Created ? new Date(i.Created).toLocaleString() : '',
                User: i.Author ? i.Author.Title : '',
                Module: i.Module,
                Page: i.Page,
                ItemId: i.ItemId,
                Severity: i.Severity,
                Error: i.Error
            }));
            const ws = XLSX.utils.json_to_sheet(cleanData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Logs");
            XLSX.writeFile(wb, `Logs_${this.props.currentListTitle}.xlsx`);
        } catch (error: any) {
            LoggerService.log('Power DataSyncWeb Part', 'LogViewer-exportLogs', 'High', 'Config', error.message || JSON.stringify(error));
        }
    }

    public render() {
        const { items, loading, selectionCount } = this.state;
        const columns: IColumn[] = [
            {
                key: 'col0',
                name: 'Date',
                fieldName: 'Created',
                minWidth: 100,
                maxWidth: 140,
                onRender: (item) => {
                    try {
                        return new Date(item.Created).toLocaleString();
                    } catch (e) {
                        return '-';
                    }
                }
            },
            { key: 'col1', name: 'Module', fieldName: 'Module', minWidth: 70, maxWidth: 100 },
            {
                key: 'col2', name: 'Severity', fieldName: 'Severity', minWidth: 60, maxWidth: 80, onRender: (item) => {
                    const color = item.Severity === 'High' ? 'red' : item.Severity === 'Medium' ? 'orange' : 'green';
                    return <span style={{ color, fontWeight: 'bold' }}>{item.Severity}</span>;
                }
            },
            {
                key: 'colErrorId',
                name: 'Error ID',
                fieldName: 'ErrorId',
                minWidth: 120,
                maxWidth: 160,
                isResizable: true,
                onRender: (item) => <span style={{ background: '#f3f2f1', padding: '2px 4px' }}>{item.ErrorId || '-'}</span>
            },
            { key: 'col3', name: 'Page', fieldName: 'Page', minWidth: 50, maxWidth: 100 },
            { key: 'col4', name: 'Error', fieldName: 'Error', minWidth: 200, maxWidth: 500, isMultiline: true }
        ];

        return (
            <Panel
                isOpen={this.props.isOpen}
                onDismiss={this.props.onDismiss}
                type={PanelType.large}
                headerText="System Logs"
                closeButtonAriaLabel="Close"
            >
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 15, gap: 10 }}>
                    <SearchBox
                        placeholder="Search logs..."
                        onSearch={this.onSearch}
                        onChange={(ev, val) => this.onSearch(val || '')} // FluentUI signature update
                        style={{ flex: 1 }} 
                    />
                    <button
                        onClick={this.exportLogs}
                        style={{ padding: '0 20px', background: '#0078d4', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}
                    >
                        {selectionCount > 0 ? `Export Selected (${selectionCount})` : 'Export All'}
                    </button>
                </div>
                
                {loading ? (
                    <Spinner size={SpinnerSize.large} label="Loading logs..." />
                ) : (
                    <div style={{ height: '70vh', overflowY: 'auto', border: '1px solid #eee' }}>
                        <DetailsList
                            items={items}
                            columns={columns}
                            selection={this.selection}
                            selectionMode={SelectionMode.multiple}
                            checkboxVisibility={CheckboxVisibility.always}
                            layoutMode={DetailsListLayoutMode.justified} // Updated to enum
                        />
                        {items.length === 0 && <div style={{ padding: 20, textAlign: 'center' }}>No logs found.</div>}
                    </div>
                )}
            </Panel>
        );
    }
}