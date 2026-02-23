import * as React from 'react';
import { useState, useCallback } from 'react';
import styles from './SpSmartSearch.module.scss';
import { ISpSmartSearchProps, IColumnConfig } from './ISpSmartSearchProps';
import { ISpSmartItem, ISearchResult, DisplayMode } from '../models/ISpSmartItem';
import { SpSmartSearchService } from '../services/SpSmartSearchService';

type ActiveTab = 'documents' | 'listItems' | 'all';

const FILE_ICONS: Record<string, string> = {
  docx: '📄', doc: '📄',
  xlsx: '📊', xls: '📊',
  pptx: '📑', ppt: '📑',
  pdf: '📕', txt: '📃',
  png: '🖼', jpg: '🖼', jpeg: '🖼',
  zip: '🗜', msg: '📧',
  default: '📎'
};

function getFileIcon(fileType?: string): string {
  if (!fileType) return '📋';
  return FILE_ICONS[fileType.toLowerCase()] || FILE_ICONS.default;
}

function formatDate(dateStr?: string): string {
  if (!dateStr) return '';
  try {
    return new Date(dateStr).toLocaleDateString();
  } catch (e) {
    return '';
  }
}

function renderTitle(item: ISpSmartItem, config: IColumnConfig): any {
  if (!config.showTitle) { return null; }
  var titleDisplay = (item.title && item.title.trim() !== '') ? item.title : 'Untitled';
  return React.createElement(
    'a',
    {
      href: item.url,
      target: '_blank',
      rel: 'noopener noreferrer',
      className: styles.resultTitle,
      title: titleDisplay
    },
    titleDisplay
  );
}

function ResultItemRow(props: { item: ISpSmartItem; config: IColumnConfig }): JSX.Element {
  var item = props.item;
  var config = props.config;
  return (
    <li className={styles.resultItem}>
      <div className={styles.resultIcon}>
        {item.type === 'listItem' ? '📋' : getFileIcon(item.fileType)}
      </div>
      <div className={styles.resultContent}>
        {renderTitle(item, config)}
        <div className={styles.resultMeta}>
          {config.showFileType && item.fileType ? <span>{'📄 ' + item.fileType.toUpperCase()}</span> : null}
          {config.showListName && item.listName ? <span>{'📂 ' + item.listName}</span> : null}
          {config.showAuthor && item.author ? <span>{'👤 ' + item.author}</span> : null}
          {config.showModifiedDate && item.modifiedDate ? <span>{'📅 ' + formatDate(item.modifiedDate)}</span> : null}
        </div>
        <div className={styles.resultMeta}>
          {config.showBusinessArea && item.businessArea ? <span>{'🏢 ' + item.businessArea}</span> : null}
          {config.showSegment && item.segment ? <span>{'🔖 ' + item.segment}</span> : null}
          {config.showProject && item.project ? <span>{'📋 ' + item.project}</span> : null}
          {config.showEstimator && item.estimator ? <span>{'👷 ' + item.estimator}</span> : null}
          {config.showBid2WinId && item.bid2WinId ? <span>{'🆔 ' + item.bid2WinId}</span> : null}
          {config.showBidDate && item.bidDate ? <span>{'📆 Bid: ' + formatDate(item.bidDate)}</span> : null}
        </div>
        <div className={styles.resultMeta}>
          {config.showCity && item.city ? <span>{'🏙 ' + item.city}</span> : null}
          {config.showCounty && item.county ? <span>{'🗺 ' + item.county}</span> : null}
          {config.showState && item.state ? <span>{'📍 ' + item.state}</span> : null}
          {config.showZipCode && item.zipCode ? <span>{'📮 ' + item.zipCode}</span> : null}
          {config.showOwner && item.owner ? <span>{'🏠 ' + item.owner}</span> : null}
        </div>
        <div className={styles.resultMeta}>
          {config.showSqYards && item.sqYards ? <span>{'📐 ' + item.sqYards + ' sq yds'}</span> : null}
          {config.showLaneMiles && item.laneMiles ? <span>{'🛣 ' + item.laneMiles + ' lane mi'}</span> : null}
          {config.showNumberOfLots && item.numberOfLots ? <span>{'🔢 ' + item.numberOfLots + ' lots'}</span> : null}
        </div>
        {item.summary ? (
          <div
            className={styles.resultSummary}
            dangerouslySetInnerHTML={{ __html: item.summary }}
          />
        ) : null}
      </div>
    </li>
  );
}

function ResultsList(props: { items: ISpSmartItem[]; emptyMessage: string; config: IColumnConfig }): JSX.Element {
  if (props.items.length === 0) {
    return (
      <div className={styles.noResults}>
        <div className={styles.noResultsIcon}>🔍</div>
        <span>{props.emptyMessage}</span>
      </div>
    );
  }
  return (
    <ul className={styles.resultsList}>
      {props.items.map(function(item) {
        return <ResultItemRow key={item.id} item={item} config={props.config} />;
      })}
    </ul>
  );
}

export const SpSmartSearch: React.FC<ISpSmartSearchProps> = (props) => {
  var context = props.context;
  var displayMode = props.displayMode;
  var searchScope = props.searchScope;
  var scopeUrl = props.scopeUrl;
  var columnConfig = props.columnConfig;

  const [searchText, setSearchText] = useState<string>('');
  const [includeDocuments, setIncludeDocuments] = useState<boolean>(displayMode !== 'listItems');
  const [includeListItems, setIncludeListItems] = useState<boolean>(displayMode !== 'documents');
  const [activeTab, setActiveTab] = useState<ActiveTab>('all');
  const [results, setResults] = useState<ISearchResult | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [hasSearched, setHasSearched] = useState<boolean>(false);

  var effectiveDisplayMode: DisplayMode =
    includeDocuments && includeListItems ? 'both' :
    includeDocuments ? 'documents' :
    includeListItems ? 'listItems' : 'both';

  const handleSearch = useCallback(async () => {
    if (!searchText.trim()) return;
    if (!scopeUrl) {
      setErrorMessage('Please configure a site or list URL in the web part properties.');
      return;
    }
    setIsLoading(true);
    setErrorMessage('');
    setHasSearched(true);
    try {
      var service = new SpSmartSearchService(context);
      var searchResults = await service.search(
        searchText.trim(), effectiveDisplayMode, searchScope, scopeUrl
      );
      setResults(searchResults);
      if (effectiveDisplayMode === 'both') {
        setActiveTab('all');
      } else if (effectiveDisplayMode === 'documents') {
        setActiveTab('documents');
      } else {
        setActiveTab('listItems');
      }
    } catch (err: any) {
      setErrorMessage(err && err.message ? err.message : 'An unexpected error occurred.');
      setResults(null);
    } finally {
      setIsLoading(false);
    }
  }, [searchText, scopeUrl, searchScope, effectiveDisplayMode, context]);

  function handleKeyDown(e: React.KeyboardEvent<HTMLInputElement>): void {
    if (e.key === 'Enter') {
      handleSearch().catch(function(err) { console.error(err); });
    }
  }

  if (!scopeUrl) {
    return (
      <div className={styles.configRequired}>
        <h3>Configuration Required</h3>
        <p>Open the web part properties panel and provide a Site URL or List/Library URL.</p>
      </div>
    );
  }

  var allItems: ISpSmartItem[] = results
    ? (results.documents || []).concat(results.listItems || [])
    : [];

  var visibleItems: ISpSmartItem[] =
    activeTab === 'documents' ? (results ? results.documents : []) :
    activeTab === 'listItems' ? (results ? results.listItems : []) :
    allItems;

  var showTabs = effectiveDisplayMode === 'both';
  var docCount = results && results.documents ? results.documents.length : 0;
  var listCount = results && results.listItems ? results.listItems.length : 0;

  return (
    <div className={styles.spSmartSearch}>
      <div className={styles.searchBar}>
        <input
          type="text"
          className={styles.searchInput}
          placeholder="Enter search terms..."
          value={searchText}
          onChange={function(e) { setSearchText(e.target.value); }}
          onKeyDown={handleKeyDown}
          aria-label="Search query"
        />
        <div className={styles.checkboxGroup}>
          {displayMode !== 'listItems' ? (
            <label>
              <input
                type="checkbox"
                checked={includeDocuments}
                onChange={function(e) { setIncludeDocuments(e.target.checked); }}
              />
              Documents
            </label>
          ) : null}
          {displayMode !== 'documents' ? (
            <label>
              <input
                type="checkbox"
                checked={includeListItems}
                onChange={function(e) { setIncludeListItems(e.target.checked); }}
              />
              List Items
            </label>
          ) : null}
          {displayMode === 'both' ? (
            <label>
              <input
                type="checkbox"
                checked={includeDocuments && includeListItems}
                onChange={function(e) {
                  setIncludeDocuments(e.target.checked);
                  setIncludeListItems(e.target.checked);
                }}
              />
              Both
            </label>
          ) : null}
        </div>
        <button
          className={styles.searchButton}
          onClick={function() { handleSearch().catch(function(err) { console.error(err); }); }}
          disabled={isLoading || !searchText.trim()}
          aria-label="Search"
        >
          {isLoading ? 'Searching...' : 'Search'}
        </button>
      </div>

      {errorMessage ? (
        <div className={styles.errorMessage} role="alert">{errorMessage}</div>
      ) : null}

      {showTabs && results ? (
        <div className={styles.tabContainer} role="tablist">
          <button
            role="tab"
            className={styles.tab + (activeTab === 'all' ? ' ' + styles.activeTab : '')}
            onClick={function() { setActiveTab('all'); }}
          >
            {'All (' + allItems.length + ')'}
          </button>
          <button
            role="tab"
            className={styles.tab + (activeTab === 'documents' ? ' ' + styles.activeTab : '')}
            onClick={function() { setActiveTab('documents'); }}
          >
            {'Documents (' + docCount + ')'}
          </button>
          <button
            role="tab"
            className={styles.tab + (activeTab === 'listItems' ? ' ' + styles.activeTab : '')}
            onClick={function() { setActiveTab('listItems'); }}
          >
            {'List Items (' + listCount + ')'}
          </button>
        </div>
      ) : null}

      <div className={styles.resultsContainer}>
        {isLoading ? (
          <div className={styles.loadingSpinner} aria-live="polite">Searching...</div>
        ) : null}
        {!isLoading && !hasSearched ? (
          <div className={styles.emptyState}>
            Enter a search term and press Search to get started.
          </div>
        ) : null}
        {!isLoading && hasSearched && results ? (
          <ResultsList
            items={visibleItems}
            config={columnConfig}
            emptyMessage={'No results found for "' + searchText + '"'}
          />
        ) : null}
      </div>
    </div>
  );
};

export default SpSmartSearch;
