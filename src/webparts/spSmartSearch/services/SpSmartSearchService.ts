import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISalesFunnelItem, ISearchResult, DisplayMode } from '../models/ISalesFunnelItem';

export class SalesFunnelSearchService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Main search entry point.
   * Routes to site-wide search or scoped URL search based on configuration.
   */
  public async search(
    query: string,
    displayMode: DisplayMode,
    searchScope: 'site' | 'url',
    scopeUrl: string
  ): Promise<ISearchResult> {
    if (!query || !scopeUrl) {
      return { documents: [], listItems: [] };
    }

    if (searchScope === 'url') {
      return this.searchByUrl(query, displayMode, scopeUrl);
    } else {
      return this.searchBySite(query, displayMode, scopeUrl);
    }
  }

  /**
   * Search all lists and document libraries within a given site URL
   */
  private async searchBySite(
    query: string,
    displayMode: DisplayMode,
    siteUrl: string
  ): Promise<ISearchResult> {
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
      console.error('SalesFunnelSearchService.searchBySite error:', error);
      throw error;
    }

    return results;
  }

  /**
   * Search a specific list or document library by its URL
   */
  private async searchByUrl(
    query: string,
    displayMode: DisplayMode,
    targetUrl: string
  ): Promise<ISearchResult> {
    const results: ISearchResult = { documents: [], listItems: [] };

    try {
      // Determine the site URL from the target URL
      const siteUrl = await this.getSiteUrlFromListUrl(targetUrl);
      const normalizedTargetUrl = targetUrl.replace(/\/$/, '');

      if (displayMode === 'documents' || displayMode === 'both') {
        results.documents = await this.searchDocumentsByLibraryUrl(query, siteUrl, normalizedTargetUrl);
      }
      if (displayMode === 'listItems' || displayMode === 'both') {
        results.listItems = await this.searchListItemsByListUrl(query, siteUrl, normalizedTargetUrl);
      }
    } catch (error) {
      console.error('SalesFunnelSearchService.searchByUrl error:', error);
      throw error;
    }

    return results;
  }

  /**
   * Search documents across all libraries in a site
   */
  private async searchDocuments(query: string, siteUrl: string): Promise<ISalesFunnelItem[]> {
    const encodedQuery = encodeURIComponent(query);
    // Use SharePoint Search REST API for full-text document search
    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodedQuery} AND path:"${siteUrl}"'` +
      `&selectproperties='Title,Path,FileExtension,LastModifiedTime,Author,HitHighlightedSummary'` +
      `&rowlimit=50` +
      `&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`; // Documents source

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      searchUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Search failed: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const rows = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

    return rows.map((row: any) => this.mapSearchRowToItem(row, 'document'));
  }

  /**
   * Search list items across all lists in a site
   */
  private async searchListItems(query: string, siteUrl: string): Promise<ISalesFunnelItem[]> {
    const encodedQuery = encodeURIComponent(query);
    // Use SharePoint Search REST API targeting list items
    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodedQuery} AND path:"${siteUrl}" AND contentclass:STS_ListItem'` +
      `&selectproperties='Title,Path,LastModifiedTime,Author,HitHighlightedSummary,ListId,SiteTitle'` +
      `&rowlimit=50`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      searchUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Search failed: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const rows = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

    return rows.map((row: any) => this.mapSearchRowToItem(row, 'listItem'));
  }

  /**
   * Search documents in a specific document library
   */
  private async searchDocumentsByLibraryUrl(
    query: string,
    siteUrl: string,
    libraryUrl: string
  ): Promise<ISalesFunnelItem[]> {
    const encodedQuery = encodeURIComponent(query);
    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodedQuery} AND path:"${libraryUrl}"'` +
      `&selectproperties='Title,Path,FileExtension,LastModifiedTime,Author,HitHighlightedSummary'` +
      `&rowlimit=50` +
      `&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      searchUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Search failed: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const rows = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

    return rows.map((row: any) => this.mapSearchRowToItem(row, 'document'));
  }

  /**
   * Search list items in a specific list
   */
  private async searchListItemsByListUrl(
    query: string,
    siteUrl: string,
    listUrl: string
  ): Promise<ISalesFunnelItem[]> {
    const encodedQuery = encodeURIComponent(query);
    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodedQuery} AND path:"${listUrl}" AND contentclass:STS_ListItem'` +
      `&selectproperties='Title,Path,LastModifiedTime,Author,HitHighlightedSummary,ListId,SiteTitle'` +
      `&rowlimit=50`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      searchUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Search failed: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const rows = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

    return rows.map((row: any) => this.mapSearchRowToItem(row, 'listItem'));
  }

  /**
   * Extract cell value from search result row by key
   */
  private getCellValue(row: any, key: string): string {
    const cell = (row.Cells || []).find((c: any) => c.Key === key);
    return cell ? cell.Value || '' : '';
  }

  /**
   * Map a raw search result row to ISalesFunnelItem
   */
  private mapSearchRowToItem(row: any, type: 'document' | 'listItem'): ISalesFunnelItem {
    const title = this.getCellValue(row, 'Title') || 'Untitled';
    const path = this.getCellValue(row, 'Path');
    const fileExt = this.getCellValue(row, 'FileExtension');
    const modified = this.getCellValue(row, 'LastModifiedTime');
    const author = this.getCellValue(row, 'Author');
    const summary = this.getCellValue(row, 'HitHighlightedSummary');
    const siteTitle = this.getCellValue(row, 'SiteTitle');

    return {
      id: path,
      title,
      type,
      url: path,
      fileType: fileExt || undefined,
      listName: siteTitle || undefined,
      siteUrl: path ? new URL(path).origin : undefined,
      modifiedDate: modified || undefined,
      author: author || undefined,
      summary: this.stripHighlightTags(summary)
    };
  }

  /**
   * Strip SharePoint hit highlight XML tags from summary
   */
  private stripHighlightTags(html: string): string {
    if (!html) return '';
    return html
      .replace(/<c0>/gi, '')
      .replace(/<\/c0>/gi, '')
      .replace(/<ddd\/>/gi, '...')
      .replace(/<[^>]+>/g, '');
  }

  /**
   * Attempt to derive the site URL from a list/library URL.
   * This handles both classic (/sites/...) and modern (/teams/...) URLs.
   */
  private async getSiteUrlFromListUrl(listUrl: string): Promise<string> {
    // Try to use the contextual site URL first
    const currentSiteUrl = this.context.pageContext.site.absoluteUrl;

    // If the list URL starts with the current site URL, use it directly
    if (listUrl.indexOf(currentSiteUrl) === 0) {
      return currentSiteUrl;
    }

    // Otherwise parse from URL — take the first 4 segments (protocol + host + /sites/sitename)
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
