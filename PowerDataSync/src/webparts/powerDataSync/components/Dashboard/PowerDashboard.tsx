import * as React from 'react';
// PnP v4 Imports
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { DataSyncHistoryService, JobRow } from '../../services/DataSyncHistoryService';
import { LoggerService } from '../LoggerService';
import { PowerDashboardProps } from '../IPowerDataSyncProps';
import { PowerDashboardState } from '../IPowerDataSyncState';

// Page button style
const pageBtnStyle = (disabled: boolean): React.CSSProperties => ({
  minWidth: 28,
  height: 28,
  padding: '0 6px',
  fontSize: 13,
  border: '1px solid #ccc',
  background: disabled ? '#f5f5f5' : '#fff',
  color: disabled ? '#999' : '#333',
  cursor: disabled ? 'not-allowed' : 'pointer',
  outline: 'none'
});

export default class PowerDashboard extends React.Component<PowerDashboardProps, PowerDashboardState> {
  private service!: DataSyncHistoryService;
  private sp!: SPFI; // Holds our PnP v4 instance

  constructor(props: PowerDashboardProps) {
    super(props);
    try {
      this.state = {
        jobs: [],
        loading: true,
        fTitle: '',
        fOwner: '',
        fStatus: '',
        fMinItems: undefined,
        fMinSuccess: undefined,
        fMinErrors: undefined,
        totalJobsCompleted: 0,
        totalItemsSync: 0,
        successRatio: 0,
        totalJobsInitiated: 0,
        totalJobsSuccess: 0,
        totalJobsFailed: 0,
        totalItemsSynchronized: 0,
        totalItemsFailed: 0,
        overallSuccessRatio: 0,
        fListName: '',
        page: 1,
        pageSize: 10
      };

      // 1. Initialize PnP v4 Factory
      this.sp = spfi().using(SPFx(this.props.context));

      // 2. Pass the SPFI instance to the DataSyncHistoryService
      this.service = new DataSyncHistoryService(this.sp, this.props.metricsListTitle);
    } catch (e: any) {
      void LoggerService.log('System', 'PowerDashboard-constructor', 'High', 'Init', e.message);
    }
  }

  public componentDidMount() {
    try {
      void this.loadData();
    } catch (e: any) {
      void LoggerService.log('System', 'componentDidMount', 'High', 'Init', e.message);
    }
  }

  private loadData = async () => {
    if (!this.props.metricsListTitle) {
      console.warn('PowerDashboard: No Metrics List Title configured.');
      this.setState({ loading: false });
      return;
    }

    try {
      this.setState({ loading: true });
      const data = await this.service.getJobs(2000);

      let completed = 0;
      let itemsSync = 0;
      let jobsInit = 0;
      let jobsSuccess = 0;
      let jobsFailed = 0;
      let itemsSuccess = 0;
      let itemsFailed = 0;

      data.forEach(j => {
        jobsInit++;
        const s = (j.Status || '').toLowerCase();
        if (s === 'completed' || s === 'completed with errors') {
          completed++;
          jobsSuccess++;
        } else if (s === 'failed') {
          jobsFailed++;
        }

        const suc = j.SuccessCount || 0;
        const fail = j.FailureCount || 0;
        itemsSync += suc;
        itemsSuccess += suc;
        itemsFailed += fail;
      });

      const ratio = jobsInit > 0 ? Math.round((jobsSuccess / jobsInit) * 100) : 0;
      const itemRatio = (itemsSuccess + itemsFailed) > 0 ? Math.round((itemsSuccess / (itemsSuccess + itemsFailed)) * 100) : 0;

      this.setState({
        jobs: data,
        loading: false,
        totalJobsCompleted: completed,
        totalItemsSync: itemsSync,
        successRatio: ratio,
        totalJobsInitiated: jobsInit,
        totalJobsSuccess: jobsSuccess,
        totalJobsFailed: jobsFailed,
        totalItemsSynchronized: itemsSuccess,
        totalItemsFailed: itemsFailed,
        overallSuccessRatio: itemRatio
      });
    } catch (err: any) {
      console.error(err);
      const logTarget = this.props.metricsListTitle || 'System';
      void LoggerService.log(logTarget, 'loadData', 'High', 'Read', err.message);
      this.setState({ loading: false });
    }
  }

  private clearFilters = () => {
    try {
      this.setState({
        fTitle: '',
        fOwner: '',
        fStatus: '',
        fListName: '',           
        fMinItems: undefined,
        fMinSuccess: undefined,
        fMinErrors: undefined,
        page: 1
      });
    } catch (e: any) {
      void LoggerService.log('System', 'clearFilters', 'Low', 'UI', e.message);
    }
  }

  // Filter
  private matches = (row: JobRow): boolean => {
    try {
      const s = this.state;
      const title = (s.fTitle || '').toLowerCase();
      const owner = (s.fOwner || '').toLowerCase();
      const listName = (s.fListName || '').toLowerCase();
      const items = (row.SuccessCount || 0) + (row.FailureCount || 0);

      if (title && row.Title && row.Title.toLowerCase().indexOf(title) === -1) return false;
      if (s.fStatus && (row.Status || '') !== s.fStatus) return false;
      if (owner && row.Owner && row.Owner.toLowerCase().indexOf(owner) === -1) return false;
      if (listName && !row.lists.some(l => l.toLowerCase().indexOf(listName) !== -1)) return false;
      if (s.fMinItems && items < s.fMinItems) return false;
      if (s.fMinSuccess && (row.SuccessCount || 0) < s.fMinSuccess) return false;
      if (s.fMinErrors && (row.FailureCount || 0) < s.fMinErrors) return false;
      
      return true;
    } catch (e: any) {
      void LoggerService.log('System', 'matches', 'Medium', 'Filter', e.message);
      return false;
    }
  }

  // Utility
  private pretty = (iso?: string): string => {
    try {
      if (!iso) return '—';
      const date = new Date(iso);
      const now = new Date();
      const diffMs = now.getTime() - date.getTime();
      if (diffMs < 0) {
        const day = ('0' + date.getDate().toString()).slice(-2);
        const month = date.toLocaleString('default', { month: 'short' });
        const year = date.getFullYear();
        return `${day} ${month} ${year}`;
      }
      const diffSec = Math.floor(diffMs / 1000);
      const diffMin = Math.floor(diffSec / 60);
      const diffHour = Math.floor(diffMin / 60);
      const diffDay = Math.floor(diffHour / 24);
      if (diffSec < 60) {
        return 'few seconds ago';
      } else if (diffMin < 60) {
        return `${diffMin} minute${diffMin > 1 ? 's' : ''} ago`;
      } else if (diffHour < 24) {
        return `${diffHour} hour${diffHour > 1 ? 's' : ''} ago`;
      } else if (diffDay === 1) {
        return 'yesterday';
      } else if (diffDay < 7) {
        const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        return `last ${days[date.getDay()]}`;
      } else {
        const day = date.getDate().toString().slice(-2);
        const month = date.toLocaleString('default', { month: 'short' });
        const year = date.getFullYear();
        return `${day} ${month} ${year}`;
      }
    } catch (e: any) {
      void LoggerService.log('System', 'pretty', 'Low', 'Format', e.message);
      return '—';
    }
  }

  public render() {
    const {
      loading,
      jobs,
      totalJobsInitiated,
      totalJobsSuccess,
      totalJobsFailed,
      overallSuccessRatio,
      totalItemsSynchronized,
      page,
      pageSize
    } = this.state;

    const filtered = jobs.filter(j => this.matches(j));
    const startIdx = (page - 1) * pageSize;
    const endIdx = startIdx + pageSize;
    const pageItems = filtered.slice(startIdx, endIdx);

    // --- STYLES ---
    const dashboardContainer: React.CSSProperties = {
      background: '#faf9f8',
      minHeight: '100%'
    };

    const headerStyle: React.CSSProperties = {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: 24,
      padding: '0 4px'
    };

    const kpiGridStyle: React.CSSProperties = {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
      gap: 16,
      marginBottom: 32
    };

    const cardStyle: React.CSSProperties = {
      background: '#fff',
      borderRadius: 4,
      padding: '20px 24px',
      boxShadow: '0 1.6px 3.6px 0 rgba(0,0,0,0.132), 0 0.3px 0.9px 0 rgba(0,0,0,0.108)',
      borderLeft: '4px solid transparent'
    };

    const kpiLabelStyle: React.CSSProperties = {
      fontSize: 12,
      fontWeight: 600,
      color: '#605e5c',
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
      marginBottom: 8
    };

    const kpiValueStyle: React.CSSProperties = {
      fontSize: 32,
      fontWeight: 300,
      color: '#323130'
    };

    return (
      <div style={dashboardContainer}>
        {/* Header */}
        <div style={headerStyle}>
          <div>
            <h2 style={{ margin: 0, fontWeight: 600, fontSize: 24, color: '#323130' }}>Data Synchronization Dashboard</h2>
            <p style={{ margin: '4px 0 0 0', color: '#605e5c', fontSize: 13 }}>Overview of recent Power DataSync jobs</p>
          </div>
          <button
            onClick={this.props.onNewJob}
            style={{
              background: '#0078d4', color: 'white', border: 'none', padding: '10px 24px',
              borderRadius: 2, fontWeight: 600, cursor: 'pointer', fontSize: 14,
              boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
            }}
          >
            + New Job
          </button>
        </div>

        {/* KPIs */}
        <div style={kpiGridStyle}>
          <div style={{ ...cardStyle, borderLeftColor: '#0078d4' }}>
            <div style={kpiLabelStyle}>Jobs Initiated</div>
            <div style={kpiValueStyle}>{loading ? '-' : totalJobsInitiated}</div>
          </div>
          <div style={{ ...cardStyle, borderLeftColor: '#107c10' }}>
            <div style={kpiLabelStyle}>Success</div>
            <div style={{ ...kpiValueStyle, color: '#107c10' }}>{loading ? '-' : totalJobsSuccess}</div>
          </div>
          <div style={{ ...cardStyle, borderLeftColor: '#d13438' }}>
            <div style={kpiLabelStyle}>Failed</div>
            <div style={{ ...kpiValueStyle, color: '#d13438' }}>{loading ? '-' : totalJobsFailed}</div>
          </div>
          <div style={{ ...cardStyle, borderLeftColor: '#00bcf2' }}>
            <div style={kpiLabelStyle}>Items Synced</div>
            <div style={kpiValueStyle}>{loading ? '-' : totalItemsSynchronized.toLocaleString()}</div>
          </div>
          <div style={{ ...cardStyle, borderLeftColor: overallSuccessRatio > 80 ? '#107c10' : '#d13438' }}>
            <div style={kpiLabelStyle}>Success Ratio</div>
            <div style={kpiValueStyle}>{loading ? '-' : `${overallSuccessRatio.toFixed(1)}%`}</div>
          </div>
        </div>

        {/* Filters & Table Section */}
        <div style={{ background: '#fff', borderRadius: 4, boxShadow: '0 1.6px 3.6px 0 rgba(0,0,0,0.132)' }}>
          {/* Toolbar */}
          <div style={{ padding: '16px 20px', borderBottom: '1px solid #edebe9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 600 }}>Import History</h3>
            <div style={{ display: 'flex', gap: 10 }}>
              <input
                placeholder="Search Title..."
                onChange={(e: any) => this.setState({ fTitle: e.target.value })}
                value={this.state.fTitle}
                style={{ padding: '6px 10px', border: '1px solid #8a8886', borderRadius: 2 }}
              />
              <button
                onClick={this.clearFilters}
                style={{ background: 'transparent', border: 'none', color: '#0078d4', cursor: 'pointer', fontWeight: 600 }}
              >
                Clear Filters
              </button>
            </div>
          </div>

          {/* Table */}
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 14 }}>
              <thead style={{ background: '#f3f2f1', color: '#605e5c' }}>
                <tr>
                  <th style={{ padding: '12px 20px', textAlign: 'left', fontWeight: 600 }}>Job Title</th>
                  <th style={{ padding: '12px 20px', textAlign: 'left', fontWeight: 600 }}>Date</th>
                  <th style={{ padding: '12px 20px', textAlign: 'left', fontWeight: 600 }}>Owner</th>
                  <th style={{ padding: '12px 20px', textAlign: 'center', fontWeight: 600 }}>Status</th>
                  <th style={{ padding: '12px 20px', textAlign: 'right', fontWeight: 600 }}>Items</th>
                  <th style={{ padding: '12px 20px', textAlign: 'right', fontWeight: 600 }}>Success</th>
                  <th style={{ padding: '12px 20px', textAlign: 'right', fontWeight: 600 }}>Errors</th>
                  <th style={{ padding: '12px 20px', textAlign: 'center', fontWeight: 600 }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 ? (
                  <tr><td colSpan={8} style={{ padding: 40, textAlign: 'center', color: '#605e5c' }}>No matching records found</td></tr>
                ) : (
                  pageItems.map(row => {
                    const statusColor = row.Status === 'Completed' ? '#107c10' : row.Status === 'Failed' ? '#d13438' : '#0078d4';
                    return (
                      <tr key={row.Id} style={{ borderBottom: '1px solid #edebe9' }}>
                        <td style={{ padding: '12px 20px', fontWeight: 600 }}>{row.Title}</td>
                        <td style={{ padding: '12px 20px', color: '#605e5c' }}>{this.pretty(row.JobStartTime)}</td>
                        <td style={{ padding: '12px 20px' }}>{row.Owner}</td>
                        <td style={{ padding: '12px 20px', textAlign: 'center' }}>
                          <span style={{
                            background: statusColor, color: 'white',
                            padding: '2px 8px', borderRadius: 12, fontSize: 11, fontWeight: 600
                          }}>
                            {row.Status}
                          </span>
                        </td>
                        <td style={{ padding: '12px 20px', textAlign: 'right' }}>{(row.SuccessCount || 0) + (row.FailureCount || 0)}</td>
                        <td style={{ padding: '12px 20px', textAlign: 'right', color: '#107c10' }}>{row.SuccessCount || 0}</td>
                        <td style={{ padding: '12px 20px', textAlign: 'right', color: '#d13438' }}>{row.FailureCount || 0}</td>
                        <td style={{ padding: '12px 20px', textAlign: 'center' }}>
                          {(row.Status === 'Failed' || row.Status === 'Completed with errors') && this.props.onResumeJob &&
                            <button
                              onClick={() => this.props.onResumeJob!(row.Id)}
                              style={{ border: '1px solid #8a8886', background: 'white', borderRadius: 2, padding: '4px 8px', cursor: 'pointer', fontSize: 12 }}
                            >
                              Resume
                            </button>
                          }
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Footer / Pagination */}
          <div style={{ padding: '12px 20px', borderTop: '1px solid #edebe9', display: 'flex', justifyContent: 'flex-end' }}>
            <div style={{ fontSize: 13, color: '#605e5c' }}>Showing {startIdx + 1} to {Math.min(endIdx, filtered.length)} of {filtered.length}</div>
          </div>
        </div>
      </div>
    );
  }
}