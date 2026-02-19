import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISpSmartItem, ISearchResult, DisplayMode } from '../models/ISpSmartItem';

export class SpSmartSearchService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async search(
    query: string,
    displayMode: DisplayMode,
    searchScope: 'site' | 'url',
    scopeUrl: string
  ): Promise<ISearchResult> {
    if (!query || !scopeUrl) return { documents: [], listItems: [] };
    if (searchScope === 'url') return this.searchByUrl(query, displayMode, scopeUrl);
    return this.searchBySite(query, displayMode, scopeUrl);
  }

  private async searchBySite(query: string, displayMode: DisplayMode, siteUrl: string): Promise<ISearchResult> {
    const results: ISearchResult = { documents: [], listItems: [] };
    const normalizedSiteUrl = siteUrl.replace(/\/$/, '');
    try {
      if (displayMode === 'documents' || displayMode === 'both') {
        results.documents = await this.searchDocuments(query, normalizedSiteUrl);
      }
      if (displayMode === 'listItems' || displayMode === 'both') {
        results.listItems = await this.searchListItems(query, normalizedSiteUrl);
      }
    } catch (error) {
      console.error('SpSmartSearchService.searchBySite error:', error);
      throw error;
    }
    return results;
  }

  private async searchByUrl(query: string, displayMode: DisplayMode, targetUrl: string): Promise<ISearchResult> {
    const results: ISearchResult = { documents: [], listItems: [] };
    try {
      const siteUrl = this.getSiteUrlFromListUrl(targetUrl);
      const normalizedTargetUrl = targetUrl.replace(/\/$/, '');
      if (displayMode === 'documents' || displayMode === 'both') {
        results.documents = await this.searchDocumentsByLibraryUrl(query, siteUrl, normalizedTargetUrl);
      }
      if (displayMode === 'listItems' || displayMode === 'both') {
        results.listItems = await this.searchListItemsByListUrl(query, siteUrl, normalizedTargetUrl);
      }
    } catch (error) {
      console.error('SpSmartSearchService.searchByUrl error:', error);
      throw error;
    }
    return results;
  }

  private async searchDocuments(query: string, siteUrl: string): Promise<ISpSmartItem[]> {
    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(query)} AND Path=${siteUrl}'` +
      `&selectproperties='Title,Path,FileExtension,LastModifiedTime,Author,HitHighlightedSummary'` +
      `&rowlimit=50` +
      `&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`;
    return this.executeSearch(searchUrl, 'document');
  }

  private async searchListItems(query: string, siteUrl: string): Promise<ISpSmartItem[]> {
    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(query)} AND Path=${siteUrl} AND contentclass:STS_ListItem'` +
      `&selectproperties='Title,Path,LastModifiedTime,Author,HitHighlightedSummary,SiteTitle'` +
      `&rowlimit=50`;
    return this.executeSearch(searchUrl, 'listItem');
  }

  private async searchDocumentsByLibraryUrl(query: string, siteUrl: string, libraryUrl: string): Promise<ISpSmartItem[]> {
    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(query)} AND Path=${libraryUrl}'` +
      `&selectproperties='Title,Path,FileExtension,LastModifiedTime,Author,HitHighlightedSummary'` +
      `&rowlimit=50` +
      `&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`;
    return this.executeSearch(searchUrl, 'document');
  }

  private async searchListItemsByListUrl(query: string, siteUrl: string, listUrl: string): Promise<ISpSmartItem[]> {
    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(query)} AND Path=${listUrl} AND contentclass:STS_ListItem'` +
      `&selectproperties='Title,Path,LastModifiedTime,Author,HitHighlightedSummary,SiteTitle'` +
      `&rowlimit=50`;
    return this.executeSearch(searchUrl, 'listItem');
  }

  private async executeSearch(searchUrl: string, type: 'document' | 'listItem'): Promise<ISpSmartItem[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      searchUrl, SPHttpClient.configurations.v1
    );
    if (!response.ok) throw new Error(`Search failed: ${response.status} ${response.statusText}`);
    const data = await response.json();
    const rows = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
    return rows.map((row: any) => this.mapRowToItem(row, type));
  }

  private getCellValue(row: any, key: string): string {
    const cell = (row.Cells || []).find((c: any) => c.Key === key);
    return cell ? cell.Value || '' : '';
  }

  private mapRowToItem(row: any, type: 'document' | 'listItem'): ISpSmartItem {
    const path = this.getCellValue(row, 'Path');
    return {
      id: path,
      title: this.getCellValue(row, 'Title') || 'Untitled',
      type,
      url: path,
      fileType: this.getCellValue(row, 'FileExtension') || undefined,
      listName: this.getCellValue(row, 'SiteTitle') || undefined,
      modifiedDate: this.getCellValue(row, 'LastModifiedTime') || undefined,
      author: this.getCellValue(row, 'Author') || undefined,
      summary: this.stripHighlightTags(this.getCellValue(row, 'HitHighlightedSummary'))
    };
  }

  private stripHighlightTags(html: string): string {
    if (!html) return '';
    return html
      .replace(/<c0>/gi, '')
      .replace(/<\/c0>/gi, '')
      .replace(/<ddd\/>/gi, '...')
      .replace(/<[^>]+>/g, '');
  }

  private getSiteUrlFromListUrl(listUrl: string): string {
    const currentSiteUrl = this.context.pageContext.site.absoluteUrl;
    if (listUrl.indexOf(currentSiteUrl) === 0) return currentSiteUrl;
    try {
      const url = new URL(listUrl);
      const segments = url.pathname.split('/').filter(Boolean);
      if (segments.length >= 2 && (segments[0] === 'sites' || segments[0] === 'teams')) {
        return `${url.origin}/${segments[0]}/${segments[1]}`;
      }
      return url.origin;
    } catch {
      return currentSiteUrl;
    }
  }
}
