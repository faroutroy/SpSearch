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

  private async searchBySite(
    query: string,
    displayMode: DisplayMode,
    siteUrl: string
  ): Promise<ISearchResult> {
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

    const baseProps: string[] = [
      'Title', 'Path', 'FileExtension', 'LastModifiedTime', 'Author',
      'Created', 'CreatedBy', 'HitHighlightedSummary', 'SiteTitle',
      'ListItemID', 'ListID'
    ];

    // Add user-specified title column if not already included
    if (this.titleColumn && this.titleColumn.trim() !== '') {
      var tc = this.titleColumn.trim();
      if (baseProps.indexOf(tc) === -1) {
        baseProps.push(tc);
      }
    }

    var contentTypeFilter = type === 'document'
      ? 'ContentTypeId:0x0101*'
      : 'ContentTypeId:0x0* NOT ContentTypeId:0x0101*';

    var pathScope = scopedUrl
      ? '(path:"' + scopedUrl + '" OR ParentLink:"' + scopedUrl + '*")'
      : 'path:"' + siteUrl + '"';

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

  // Returns value as string — handles null ValueType correctly
  private getCellValue(row: any, key: string): string {
    const cells = row.Cells && row.Cells.results ? row.Cells.results : (row.Cells || []);
    const cell = cells.find((c: any) => c.Key === key);
    if (!cell) return '';
    if (cell.ValueType === 'Null' || cell.Value === null || cell.Value === undefined) return '';
    return String(cell.Value);
  }

  // Stores ALL keys — value is empty string if null, so keys are always present
  private getAllCells(row: any): Record<string, string> {
    const cells = row.Cells && row.Cells.results ? row.Cells.results : (row.Cells || []);
    const result: Record<string, string> = {};
    cells.forEach(function(c: any) {
      if (c.Key) {
        // Store empty string for null values — key still present for matching
        if (c.ValueType === 'Null' || c.Value === null || c.Value === undefined) {
          result[c.Key] = '';
        } else {
          result[c.Key] = String(c.Value);
        }
      }
    });
    return result;
  }

  private resolveTitleFromColumn(allCells: Record<string, string>): string {
    if (!this.titleColumn || this.titleColumn.trim() === '') return '';

    var tc = this.titleColumn.trim();
    var tcLower = tc.toLowerCase();
    var keys = Object.keys(allCells);

    // Step 1 — exact match
    if (allCells[tc] !== undefined && allCells[tc].trim() !== '') {
      return allCells[tc];
    }

    // Step 2 — case-insensitive exact match
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].toLowerCase() === tcLower) {
        if (allCells[keys[i]] && allCells[keys[i]].trim() !== '') {
          return allCells[keys[i]];
        }
      }
    }

    // Step 3 — partial match (e.g. "Bid2Win" matches "Bid2WinIDOWSNMBR")
    for (var j = 0; j < keys.length; j++) {
      if (keys[j].toLowerCase().indexOf(tcLower) !== -1) {
        if (allCells[keys[j]] && allCells[keys[j]].trim() !== '') {
          return allCells[keys[j]];
        }
      }
    }

    return '';
  }

  private buildItemUrl(
    path: string,
    type: 'document' | 'listItem',
    allCells: Record<string, string>
  ): string {
    if (type !== 'listItem') return path;

    var listItemId = allCells['ListItemID'] || '';
    var listId = allCells['ListID'] || '';

    if (listItemId && listId) {
      var siteAbsUrl = this.context.pageContext.site.absoluteUrl;
      return siteAbsUrl + '/_layouts/15/listform.aspx?PageType=4&ListId=' + listId + '&ID=' + listItemId;
    }

    return path;
  }

  private mapRowToItem(row: any, type: 'document' | 'listItem', searchQuery: string): ISpSmartItem {
    const path = this.getCellValue(row, 'Path');
    const allCells = this.getAllCells(row);

    // Try to get title from user-specified column
    var displayTitle = this.resolveTitleFromColumn(allCells);

    // Fallback to search keyword if column is empty or null
    if (!displayTitle || displayTitle.trim() === '') {
      displayTitle = searchQuery;
    }

    var itemUrl = this.buildItemUrl(path, type, allCells);

    return {
      id: path,
      title: displayTitle,
      type: type,
      url: itemUrl,
      fileType: allCells['FileExtension'] || undefined,
      listName: allCells['SiteTitle'] || undefined,
      modifiedDate: allCells['LastModifiedTime'] || undefined,
      author: allCells['Author'] || undefined,
      summary: this.stripHighlightTags(allCells['HitHighlightedSummary'] || ''),
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
