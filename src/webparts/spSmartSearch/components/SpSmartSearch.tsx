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

function getFileIcon(fileType: string | undefined): string {
  if (!fileType) return '📋';
  return FILE_ICONS[fileType.toLowerCase()] || FILE_ICONS.default;
}

function formatDate(dateStr: string | undefined): string {
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

function ResultItemRow(props: { item: ISpSmartItem; config: IColumnConfig }): any {
  var item = props.item;
  var config = props.config;
  return React.createElement(
    'li',
    { className: styles.resultItem },
    React.createElement(
      'div',
      { className: styles.resultIcon },
      item.type === 'listItem' ? '📋' : getFileIcon(item.fileType)
    ),
    React.createElement(
      'div',
      { className: styles.resultContent },
      renderTitle(item, config),
      React.createElement(
        'div',
        { className: styles.resultMeta },
        config.showFileType && item.fileType ? React.createElement('span', null, '📄 ' + item.fileType.toUpperCase()) : null,
        config.showListName && item.listName ? React.createElement('span', null, '📂 ' + item.listName) : null,
        config.showAuthor && item.author ? React.createElement('span', null, '👤 ' + item.author) : null,
        config.showModifiedDate && item.modifiedDate ? React.createElement('span', null, '📅 ' + formatDate(item.modifiedDate)) : null
      ),
      React.createElement(
        'div',
        { className: styles.resultMeta },
        config.showBusinessArea && item.businessArea ? React.createElement('span', null, '🏢 ' + item.businessArea) : null,
        config.showSegment && item.segment ? React.createElement('span', null, '🔖 ' + item.segment) : null,
        config.showProject && item.project ? React.createElement('span', null, '📋 ' + item.project) : null,
        config.showEstimator && item.estimator ? React.createElement('span', null, '👷 ' + item.estimator) : null,
        config.showBid2WinId && item.bid2WinId ? React.createElement('span', null, '🆔 ' + item.bid2WinId) : null,
        config.showBidDate && item.bidDate ? React.createElement('span', null, '📆 Bid: ' + formatDate(item.bidDate)) : null
      ),
      React.createElement(
        'div',
        { className: styles.resultMeta },
        config.showCity && item.city ? React.createElement('span', null, '🏙 ' + item.city) : null,
        config.showCounty && item.county ? React.createElement('span', null, '🗺 ' + item.county) : null,
        config.showState && item.state ? React.createElement('span', null, '📍 ' + item.state) : null,
        config.showZipCode && item.zipCode ? React.createElement('span', null, '📮 ' + item.zipCode) : null,
        config.showOwner && item.owner ? React.createElement('span', null, '🏠 ' + item.owner) : null
      ),
      React.createElement(
        'div',
        { className: styles.resultMeta },
        config.showSqYards && item.sqYards ? React.createElement('span', null, '📐 ' + item.sqYards + ' sq yds') : null,
        config.showLaneMiles && item.laneMiles ? React.createElement('span', null, '🛣 ' + item.laneMiles + ' lane mi') : null,
        config.showNumberOfLots && item.numberOfLots ? React.createElement('span', null, '🔢 ' + item.numberOfLots + ' lots') : null
      ),
      item.summary ? React.createElement('div', { className: styles.resultSummary, dangerouslySetInnerHTML: { __html: item.summary } }) : null
    )
  );
}

function ResultsList(props: { items: ISpSmartItem[]; emptyMessage: string; config: IColumnConfig }): any {
  if (props.items.length === 0) {
    return React.createElement(
      'div',
      { className: styles.noResults },
      React.createElement('div', { className: styles.noResultsIcon }, '🔍'),
      React.createElement('span', null, props.emptyMessage)
    );
  }
  return React.createElement(
    'ul',
    { className: styles.resultsList },
    props.items.map(function(item) {
      return React.createElement(ResultItemRow, { key: item.id, item: item, config: props.config });
    })
  );
}

export const SpSmartSearch: React.FC<ISpSmartSearchProps> = (props) => {
  var context = props.context;
  var displayMode = props.displayMode;
  var searchScope = props.searchScope;
  var scopeUrl = props.scopeUrl;
  var columnConfig = props.columnConfig;
  var titleField = props.titleField;

  const [searchText, setSearchText] = useState('');
  const [includeDocuments, setIncludeDocuments] = useState(displayMode !== 'listItems');
  const [includeListItems, setIncludeListItems] = useState(displayMode !== 'documents');
  const [activeTab, setActiveTab] = useState('all' as ActiveTab);
  const [results, setResults] = useState(null as ISearchResult | null);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [hasSearched, setHasSearched] = useState(false);

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
      var service = new SpSmartSearchService(context, titleField);
      var searchResults = await service.search(searchText.trim(), effectiveDisplayMode, searchScope, scopeUrl);
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
  }, [searchText, scopeUrl, searchScope, effectiveDisplayMode, context, titleField]);

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
              <input type="checkbox" checked={includeDocuments}
                onChange={function(e) { setIncludeDocuments(e.target.checked); }} />
              Documents
            </label>
          ) : null}
          {displayMode !== 'documents' ? (
            <label>
              <input type="checkbox" checked={includeListItems}
                onChange={function(e) { setIncludeListItems(e.target.checked); }} />
              List Items
            </label>
          ) : null}
          {displayMode === 'both' ? (
            <label>
              <input type="checkbox" checked={includeDocuments && includeListItems}
                onChange={function(e) {
                  setIncludeDocuments(e.target.checked);
                  setIncludeListItems(e.target.checked);
                }} />
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
          React.createElement(ResultsList, {
            items: visibleItems,
            config: columnConfig,
            emptyMessage: 'No results found for "' + searchText + '"'
          })
        ) : null}
      </div>
    </div>
  );
};

export default SpSmartSearch;
