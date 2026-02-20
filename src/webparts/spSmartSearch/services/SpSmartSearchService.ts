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
      console.error('SpSmartSearchService.searchBySite error:', error);
      throw error;
    }
    return results;
  }

  private async searchByUrl(
    query: string,
    displayMode: DisplayMode,
    targetUrl: string
  ): Promise<ISearchResult> {
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

  private readonly docSelectProps: string =
    'Title,Path,FileExtension,LastModifiedTime,Author,HitHighlightedSummary,SiteTitle,' +
    'BusinessArea,BidDate,Estimator,Bid2WinID,Segment,SqYards,LaneMiles,' +
    'NumberOfLots,City,County,State,ZipCode,Owner,Project';

  private readonly listSelectProps: string =
    'Title,Path,LastModifiedTime,Author,HitHighlightedSummary,SiteTitle,' +
    'BusinessArea,BidDate,Estimator,Bid2WinID,Segment,SqYards,LaneMiles,' +
    'NumberOfLots,City,County,State,ZipCode,Owner,Project';

  private async searchDocuments(query: string, siteUrl: string): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();
    const encodedSiteUrl = encodeURIComponent(`"${siteUrl}"`);

    const queryText = `(${query}*)`;
    const queryTemplate =
      `{searchTerms} ` +
      `(siteId:{${siteId}} OR siteId:${siteId}) ` +
      `(webId:{${webId}} OR webId:${webId}) ` +
      `path:"${siteUrl}" ` +
      `ContentTypeId:0x0101*`;

    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(queryText)}'` +
      `&querytemplate='${encodeURIComponent(queryTemplate)}'` +
      `&selectproperties='${this.docSelectProps}'` +
      `&rowlimit=50` +
      `&BypassResultTypes=true` +
      `&EnableQueryRules=false` +
      `&TrimDuplicates=false`;

    return this.executeSearch(searchUrl, 'document');
  }

  private async searchListItems(query: string, siteUrl: string): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();

    const queryText = `(${query}*)`;
    const queryTemplate =
      `{searchTerms} ` +
      `(siteId:{${siteId}} OR siteId:${siteId}) ` +
      `(webId:{${webId}} OR webId:${webId}) ` +
      `path:"${siteUrl}" ` +
      `ContentTypeId:0x0* ` +
      `NOT ContentTypeId:0x0101*`;

    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(queryText)}'` +
      `&querytemplate='${encodeURIComponent(queryTemplate)}'` +
      `&selectproperties='${this.listSelectProps}'` +
      `&rowlimit=50` +
      `&BypassResultTypes=true` +
      `&EnableQueryRules=false` +
      `&TrimDuplicates=false`;

    return this.executeSearch(searchUrl, 'listItem');
  }

  private async searchDocumentsByLibraryUrl(
    query: string,
    siteUrl: string,
    libraryUrl: string
  ): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();

    const queryText = `(${query}*)`;
    const queryTemplate =
      `{searchTerms} ` +
      `(siteId:{${siteId}} OR siteId:${siteId}) ` +
      `(webId:{${webId}} OR webId:${webId}) ` +
      `(path:"${libraryUrl}" OR ParentLink:"${libraryUrl}*") ` +
      `ContentTypeId:0x0101*`;

    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(queryText)}'` +
      `&querytemplate='${encodeURIComponent(queryTemplate)}'` +
      `&selectproperties='${this.docSelectProps}'` +
      `&rowlimit=50` +
      `&BypassResultTypes=true` +
      `&EnableQueryRules=false` +
      `&TrimDuplicates=false`;

    return this.executeSearch(searchUrl, 'document');
  }

  private async searchListItemsByListUrl(
    query: string,
    siteUrl: string,
    listUrl: string
  ): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();

    const queryText = `(${query}*)`;
    const queryTemplate =
      `{searchTerms} ` +
      `(siteId:{${siteId}} OR siteId:${siteId}) ` +
      `(webId:{${webId}} OR webId:${webId}) ` +
      `(path:"${listUrl}" OR ParentLink:"${listUrl}*") ` +
      `ContentTypeId:0x0* ` +
      `NOT ContentTypeId:0x0101*`;

    const searchUrl = `${siteUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(queryText)}'` +
      `&querytemplate='${encodeURIComponent(queryTemplate)}'` +
      `&selectproperties='${this.listSelectProps}'` +
      `&rowlimit=50` +
      `&BypassResultTypes=true` +
      `&EnableQueryRules=false` +
      `&TrimDuplicates=false`;

    return this.executeSearch(searchUrl, 'listItem');
  }

  private async executeSearch(
    searchUrl: string,
    type: 'document' | 'listItem'
  ): Promise<ISpSmartItem[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      searchUrl, SPHttpClient.configurations.v1
    );
    if (!response.ok) throw new Error(`Search failed: ${response.status} ${response.statusText}`);
    const data = await response.json();
    const rows =
      data &&
      data.PrimaryQueryResult &&
      data.PrimaryQueryResult.RelevantResults &&
      data.PrimaryQueryResult.RelevantResults.Table &&
      data.PrimaryQueryResult.RelevantResults.Table.Rows
        ? data.PrimaryQueryResult.RelevantResults.Table.Rows
        : [];
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
      summary: this.stripHighlightTags(this.getCellValue(row, 'HitHighlightedSummary')),
      businessArea: this.getCellValue(row, 'BusinessArea') || undefined,
      bidDate: this.getCellValue(row, 'BidDate') || undefined,
      estimator: this.getCellValue(row, 'Estimator') || undefined,
      bid2WinId: this.getCellValue(row, 'Bid2WinID') || undefined,
      segment: this.getCellValue(row, 'Segment') || undefined,
      sqYards: this.getCellValue(row, 'SqYards') || undefined,
      laneMiles: this.getCellValue(row, 'LaneMiles') || undefined,
      numberOfLots: this.getCellValue(row, 'NumberOfLots') || undefined,
      city: this.getCellValue(row, 'City') || undefined,
      county: this.getCellValue(row, 'County') || undefined,
      state: this.getCellValue(row, 'State') || undefined,
      zipCode: this.getCellValue(row, 'ZipCode') || undefined,
      owner: this.getCellValue(row, 'Owner') || undefined,
      project: this.getCellValue(row, 'Project') || undefined
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
