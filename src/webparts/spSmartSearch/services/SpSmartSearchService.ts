import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { ISpSmartItem, ISearchResult, DisplayMode } from '../models/ISpSmartItem';

export class SpSmartSearchService {
  private context: WebPartContext;
  private titleColumn: string;

  constructor(context: WebPartContext, titleColumn: string) {
    this.context = context;
    this.titleColumn = titleColumn || '';
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
        results.documents = await this.executeSearch(query, normalizedSiteUrl, 'document', null);
      }
      if (displayMode === 'listItems' || displayMode === 'both') {
        results.listItems = await this.executeSearch(query, normalizedSiteUrl, 'listItem', null);
      }
    } catch (error) {
      console.error('searchBySite error:', error);
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
        results.documents = await this.executeSearch(query, siteUrl, 'document', normalizedTargetUrl);
      }
      if (displayMode === 'listItems' || displayMode === 'both') {
        results.listItems = await this.executeSearch(query, siteUrl, 'listItem', normalizedTargetUrl);
      }
    } catch (error) {
      console.error('searchByUrl error:', error);
      throw error;
    }
    return results;
  }

  private async executeSearch(
    query: string,
    siteUrl: string,
    type: 'document' | 'listItem',
    scopedUrl: string | null
  ): Promise<ISpSmartItem[]> {
    const siteId = this.context.pageContext.site.id.toString();
    const webId = this.context.pageContext.web.id.toString();

    // Build select properties dynamically
    // Always include base properties + the user-specified title column
    const baseProps = [
      'Title', 'Path', 'FileExtension', 'LastModifiedTime', 'Author',
      'Created', 'CreatedBy', 'HitHighlightedSummary', 'SiteTitle', 'ListItemID'
    ];

    // Add title column if specified and not already in base
    if (this.titleColumn && this.titleColumn.trim() !== '') {
      var tc = this.titleColumn.trim();
      if (baseProps.indexOf(tc) === -1) {
        baseProps.push(tc);
      }
    }

    // Build content type filter
    var contentTypeFilter = '';
    if (type === 'document') {
      contentTypeFilter = 'ContentTypeId:0x0101*';
    } else {
      contentTypeFilter = 'ContentTypeId:0x0* NOT ContentTypeId:0x0101*';
    }

    // Build path scope
    var pathScope = '';
    if (scopedUrl) {
      pathScope = '(path:"' + scopedUrl + '" OR ParentLink:"' + scopedUrl + '*")';
    } else {
      pathScope = 'path:"' + siteUrl + '"';
    }

    const queryTemplate =
      '{searchTerms} ' +
      '(siteId:{' + siteId + '} OR siteId:' + siteId + ') ' +
      '(webId:{' + webId + '} OR webId:' + webId + ') ' +
      pathScope + ' ' +
      contentTypeFilter;

    const body = {
      request: {
        '__metadata': { 'type': 'Microsoft.Office.Server.Search.REST.SearchRequest' },
        'Querytext': '(' + query + '*)',
        'QueryTemplate': queryTemplate,
        'SelectProperties': {
          '__metadata': { 'type': 'Collection(Edm.String)' },
          'results': baseProps
        },
        'RowLimit': 50,
        'BypassResultTypes': true,
        'EnableQueryRules': false,
        'TrimDuplicates': false
      }
    };

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

    return rows.map((row: any) => this.mapRowToItem(row, type, query));
  }

  private getCellValue(row: any, key: string): string {
    const cells = row.Cells && row.Cells.results ? row.Cells.results : (row.Cells || []);
    const cell = cells.find((c: any) => c.Key === key);
    return cell ? cell.Value || '' : '';
  }

  private getAllCells(row: any): Record<string, string> {
    const cells = row.Cells && row.Cells.results ? row.Cells.results : (row.Cells || []);
    const result: Record<string, string> = {};
    cells.forEach((c: any) => {
      if (c.Key && c.Value) {
        result[c.Key] = c.Value;
      }
    });
    return result;
  }

  private mapRowToItem(row: any, type: 'document' | 'listItem', searchQuery: string): ISpSmartItem {
    const path = this.getCellValue(row, 'Path');
    const allCells = this.getAllCells(row);

    // Build title from user-specified column
    let displayTitle = '';

    if (this.titleColumn && this.titleColumn.trim() !== '') {
      // Try exact column name as provided
      displayTitle = this.getCellValue(row, this.titleColumn.trim());
    }

    // If still empty — use search keyword as fallback instead of "Untitled"
    if (!displayTitle || displayTitle.trim() === '') {
      displayTitle = searchQuery;
    }

    // Build the correct item URL
    // For list items — construct DispForm URL properly
    var itemUrl = path;
    if (type === 'listItem') {
      var listItemId = this.getCellValue(row, 'ListItemID');
      if (listItemId && path.toLowerCase().indexOf('dispform') !== -1) {
        // Path is already a DispForm URL — use as-is
        itemUrl = path;
      } else if (listItemId) {
        // Construct proper DispForm URL from site + list path
        var siteAbsUrl = this.context.pageContext.site.absoluteUrl;
        var listPath = this.getCellValue(row, 'ParentLink') || '';
        if (listPath) {
          itemUrl = siteAbsUrl + '/_layouts/15/listform.aspx?PageType=4&ListId=' +
            this.getCellValue(row, 'ListID') + '&ID=' + listItemId;
        } else {
          itemUrl = path;
        }
      }
    }

    return {
      id: path,
      title: displayTitle,
      type: type,
      url: itemUrl,
      fileType: this.getCellValue(row, 'FileExtension') || undefined,
      listName: this.getCellValue(row, 'SiteTitle') || undefined,
      modifiedDate: this.getCellValue(row, 'LastModifiedTime') || undefined,
      author: this.getCellValue(row, 'Author') || undefined,
      summary: this.stripHighlightTags(this.getCellValue(row, 'HitHighlightedSummary')),
      // Store all raw cells so component can access any field by name
      rawFields: allCells
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
