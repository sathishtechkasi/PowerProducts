import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface JobRow {
  Id: number;
  Title: string;
  Status?: string;
  SuccessCount: number;
  FailureCount: number;
  ItemstoImport: number;
  JobStartTime?: string;
  JobEndTime?: string;
  Owner?: string;
  lists: string[];
  logUrl?: string;
}

export class DataSyncHistoryService {
  private sp: SPFI;
  private listTitle: string;

  // We now pass the initialized SPFI instance from the web part instead of a siteUrl string
  constructor(sp: SPFI, listTitle: string) {
    this.sp = sp;
    this.listTitle = listTitle || "DataSycnHistory"; 
  }

  private getUrlFromHyperlink(field: any): string {
    if (field && typeof field === 'object' && field.Url) {
      return field.Url;
    }
    if (typeof field === 'string') {
      const parts = field.split(',').map(p => p.trim());
      return parts[0] || '';
    }
    return '';
  }

  public async getJobs(top: number = 500): Promise<JobRow[]> {
    try {
      // Notice how we use () at the end instead of .get() to execute the query in v4
      const items: any[] = await this.sp.web.lists.getByTitle(this.listTitle).items
        .select('Id,Title,DataSycnList,FailureFile,Status,SuccessCount,FailureCount,ItemstoImport,JobStartTime,JobEndTime,Author/Title')
        .orderBy('Id', false)
        .expand('File,Author')
        .top(top)(); 
        
      return items.map(it => ({
        Id: it.Id,
        Title: it.Title || '',
        lists: Array.isArray(it.DataSycnList) ? it.DataSycnList : (it.DataSycnList ? it.DataSycnList.split(',').map((l: string) => l.trim()) : []),
        Status: it.Status || 'Unknown',
        logUrl: this.getUrlFromHyperlink(it.FailureFile),
        SuccessCount: it.SuccessCount || 0,
        FailureCount: it.FailureCount || 0,
        ItemstoImport: it.ItemstoImport || 0,
        JobStartTime: it.JobStartTime,
        JobEndTime: it.JobEndTime,
        Owner: it.Author ? it.Author.Title : ''
      }));
    } catch (e) {
      // Fallback without expanding Author/File
      const items: any[] = await this.sp.web.lists.getByTitle(this.listTitle).items
        .select('Id,Title,DataSycnList,FailureFile,Status,SuccessCount,FailureCount,ItemstoImport,JobStartTime,JobEndTime')
        .orderBy('Id', false)
        .expand('File')
        .top(top)();

      return items.map(it => ({
        Id: it.Id,
        Title: it.Title || '',
        lists: Array.isArray(it.DataSycnList) ? it.DataSycnList : (it.DataSycnList ? it.DataSycnList.split(',').map((l: string) => l.trim()) : []),
        Status: it.Status || 'Unknown',
        logUrl: this.getUrlFromHyperlink(it.FailureFile),
        SuccessCount: it.SuccessCount || 0,
        FailureCount: it.FailureCount || 0,
        ItemstoImport: it.ItemstoImport || 0,
        JobStartTime: it.JobStartTime,
        JobEndTime: it.JobEndTime,
        Owner: '—'
      }));
    }
  }
}