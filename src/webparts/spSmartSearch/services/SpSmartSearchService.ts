import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { ISpSmartItem, ISearchResult, DisplayMode } from '../models/ISpSmartItem';

export class SpSmartSearchService {
  private context: WebPartContext;
  private titleField: string;

  constructor(context: WebPartContext, titleField: string) {
    this.context = context;
    this.titleField = titleField || 'auto';
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
      console.error('searchBySite error:', error);
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
      console.error('searchByUrl error:', error);
      throw error;
    }
    return results;
  }

  private buildSelectProperties(props: string[]): any {
    return {
      '__metadata': { 'type': 'Collection(Edm.String)' },
      'results': props
    };
  }

  private readonly docProps: string[] = [
    'Title', 'Path', 'FileExtension', 'LastModifiedTime', 'Author',
    'HitHighlightedSummary', 'SiteTitle', 'BusinessArea', 'BidDate',
    'Estimator', 'Bid2WinID', 'Segment', 'SqYards', 'LaneMiles',
    'NumberOfLots', 'City', 'County', 'State', 'ZipCode', 'Owner', 'Project'
  ];

  private readonly listProps: string[] = [
    'Title', 'Path', 'LastModifiedTime', 'Author',
    'HitHighlightedSummary', 'SiteTitle', 'ListItemID',
    'BusinessArea', 'BidDate', 'Estimator', 'Bid2WinID', 'Segment', 'SqYards', 'LaneMiles',
    'NumberOfLots', 'City', 'County', 'State', 'ZipCode', 'Owner', 'Project'
  ];

  private async searchDocuments(query: string, siteUrl: string): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();
    const body = {
      request: {
        '__metadata': { 'type': 'Microsoft.Office.Server.Search.REST.SearchRequest' },
        'Querytext': '(' + query + '*)',
        'QueryTemplate': '{searchTerms} (siteId:{' + siteId + '} OR siteId:' + siteId + ') (webId:{' + webId + '} OR webId:' + webId + ') path:"' + siteUrl + '" ContentTypeId:0x0101*',
        'SelectProperties': this.buildSelectProperties(this.docProps),
        'RowLimit': 50,
        'BypassResultTypes': true,
        'EnableQueryRules': false,
        'TrimDuplicates': false
      }
    };
    return this.executePostSearch(siteUrl, body, 'document');
  }

  private async searchListItems(query: string, siteUrl: string): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();
    const body = {
      request: {
        '__metadata': { 'type': 'Microsoft.Office.Server.Search.REST.SearchRequest' },
        'Querytext': '(' + query + '*)',
        'QueryTemplate': '{searchTerms} (siteId:{' + siteId + '} OR siteId:' + siteId + ') (webId:{' + webId + '} OR webId:' + webId + ') path:"' + siteUrl + '" ContentTypeId:0x0* NOT ContentTypeId:0x0101*',
        'SelectProperties': this.buildSelectProperties(this.listProps),
        'RowLimit': 50,
        'BypassResultTypes': true,
        'EnableQueryRules': false,
        'TrimDuplicates': false
      }
    };
    return this.executePostSearch(siteUrl, body, 'listItem');
  }

  private async searchDocumentsByLibraryUrl(
    query: string,
    siteUrl: string,
    libraryUrl: string
  ): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();
    const body = {
      request: {
        '__metadata': { 'type': 'Microsoft.Office.Server.Search.REST.SearchRequest' },
        'Querytext': '(' + query + '*)',
        'QueryTemplate': '{searchTerms} (siteId:{' + siteId + '} OR siteId:' + siteId + ') (webId:{' + webId + '} OR webId:' + webId + ') (path:"' + libraryUrl + '" OR ParentLink:"' + libraryUrl + '*") ContentTypeId:0x0101*',
        'SelectProperties': this.buildSelectProperties(this.docProps),
        'RowLimit': 50,
        'BypassResultTypes': true,
        'EnableQueryRules': false,
        'TrimDuplicates': false
      }
    };
    return this.executePostSearch(siteUrl, body, 'document');
  }

  private async searchListItemsByListUrl(
    query: string,
    siteUrl: string,
    listUrl: string
  ): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();
    const body = {
      request: {
        '__metadata': { 'type': 'Microsoft.Office.Server.Search.REST.SearchRequest' },
        'Querytext': '(' + query + '*)',
        'QueryTemplate': '{searchTerms} (siteId:{' + siteId + '} OR siteId:' + siteId + ') (webId:{' + webId + '} OR webId:' + webId + ') (path:"' + listUrl + '" OR ParentLink:"' + listUrl + '*") ContentTypeId:0x0* NOT ContentTypeId:0x0101*',
        'SelectProperties': this.buildSelectProperties(this.listProps),
        'RowLimit': 50,
        'BypassResultTypes': true,
        'EnableQueryRules': false,
        'TrimDuplicates': false
      }
    };
    return this.executePostSearch(siteUrl, body, 'listItem');
  }

  private async executePostSearch(
    siteUrl: string,
    body: any,
    type: 'document' | 'listItem'
  ): Promise<ISpSmartItem[]> {
    const searchUrl = siteUrl + '/_api/search/postquery';
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '3.0'
      },
      body: JSON.stringify(body)
    };
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      searchUrl,
      SPHttpClient.configurations.v1,
      options
    );
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error('Search failed (' + response.status + '): ' + errorText);
    }
    const data = await response.json();
    const rows =
      data &&
      data.d &&
      data.d.postquery &&
      data.d.postquery.PrimaryQueryResult &&
      data.d.postquery.PrimaryQueryResult.RelevantResults &&
      data.d.postquery.PrimaryQueryResult.RelevantResults.Table &&
      data.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows &&
      data.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results
        ? data.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results
        : [];
    return rows.map((row: any) => this.mapRowToItem(row, type));
  }

  private getCellValue(row: any, key: string): string {
    const cells = row.Cells && row.Cells.results ? row.Cells.results : (row.Cells || []);
    const cell = cells.find((c: any) => c.Key === key);
    return cell ? cell.Value || '' : '';
  }

  private mapRowToItem(row: any, type: 'document' | 'listItem'): ISpSmartItem {
    const path = this.getCellValue(row, 'Path');
    const rawTitle = this.getCellValue(row, 'Title');
    const bid2WinId = this.getCellValue(row, 'Bid2WinID') || undefined;
    const project = this.getCellValue(row, 'Project') || undefined;
    const listItemId = this.getCellValue(row, 'ListItemID') || undefined;

    let displayTitle: string;

    if (type === 'listItem') {
      if (this.titleField && this.titleField !== 'auto') {
        const fieldVal = this.getCellValue(row, this.titleField);
        displayTitle = fieldVal && fieldVal.trim() !== '' ? fieldVal : 'Untitled';
      } else {
        const isDispForm = !rawTitle || rawTitle.toLowerCase().indexOf('dispform') !== -1;
        if (bid2WinId) {
          displayTitle = bid2WinId;
        } else if (project) {
          displayTitle = project;
        } else if (!isDispForm && rawTitle && rawTitle.trim() !== '') {
          displayTitle = rawTitle;
        } else if (listItemId) {
          displayTitle = 'Item #' + listItemId;
        } else {
          displayTitle = 'Untitled';
        }
      }
    } else {
      displayTitle = (rawTitle && rawTitle.trim() !== '') ? rawTitle : 'Untitled';
    }

    return {
      id: path,
      title: displayTitle,
      type: type,
      url: path,
      fileType: this.getCellValue(row, 'FileExtension') || undefined,
      listName: this.getCellValue(row, 'SiteTitle') || undefined,
      modifiedDate: this.getCellValue(row, 'LastModifiedTime') || undefined,
      author: this.getCellValue(row, 'Author') || undefined,
      summary: this.stripHighlightTags(this.getCellValue(row, 'HitHighlightedSummary')),
      businessArea: this.getCellValue(row, 'BusinessArea') || undefined,
      bidDate: this.getCellValue(row, 'BidDate') || undefined,
      estimator: this.getCellValue(row, 'Estimator') || undefined,
      bid2WinId: bid2WinId,
      segment: this.getCellValue(row, 'Segment') || undefined,
      sqYards: this.getCellValue(row, 'SqYards') || undefined,
      laneMiles: this.getCellValue(row, 'LaneMiles') || undefined,
      numberOfLots: this.getCellValue(row, 'NumberOfLots') || undefined,
      city: this.getCellValue(row, 'City') || undefined,
      county: this.getCellValue(row, 'County') || undefined,
      state: this.getCellValue(row, 'State') || undefined,
      zipCode: this.getCellValue(row, 'ZipCode') || undefined,
      owner: this.getCellValue(row, 'Owner') || undefined,
      project: project
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
        return url.origin + '/' + segments[0] + '/' + segments[1];
      }
      return url.origin;
    } catch (e) {
      return currentSiteUrl;
    }
  }
}
