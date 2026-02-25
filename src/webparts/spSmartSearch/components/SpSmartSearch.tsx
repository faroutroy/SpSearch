import * as React from 'react';
import { useState, useCallback } from 'react';
import styles from './SpSmartSearch.module.scss';
import { ISpSmartSearchProps } from './ISpSmartSearchProps';
import { ISpSmartItem, ISearchResult, DisplayMode } from '../models/ISpSmartItem';
import { SpSmartSearchService } from '../services/SpSmartSearchService';

type ActiveTab = 'documents' | 'listItems' | 'all';

function formatDate(val: string): string {
  if (!val) return '';
  try {
    var d = new Date(val);
    if (isNaN(d.getTime())) return val;
    return d.toLocaleDateString();
  } catch (e) {
    return val;
  }
}

function ResultItemRow(props: {
  item: ISpSmartItem;
}): any {
  var item = props.item;

  var modifiedBy = item.author || '';
  var modifiedOn = item.modifiedDate ? formatDate(item.modifiedDate) : '';

  var metaItems: any[] = [];

  if (modifiedBy) {
    metaItems.push(
      React.createElement('span', { key: 'modby', className: styles.metaItem },
        React.createElement('span', { className: styles.metaLabel }, 'Modified by: '),
        React.createElement('span', { className: styles.metaValue }, modifiedBy)
      )
    );
  }

  if (modifiedOn) {
    metaItems.push(
      React.createElement('span', { key: 'modon', className: styles.metaItem },
        React.createElement('span', { className: styles.metaLabel }, 'Modified on: '),
        React.createElement('span', { className: styles.metaValue }, modifiedOn)
      )
    );
  }

  return React.createElement(
    'li',
    { className: styles.resultItem },
    React.createElement('div', { className: styles.resultIcon },
      item.type === 'listItem' ? '📋' : '📄'
    ),
    React.createElement('div', { className: styles.resultContent },
      React.createElement('a', {
        href: item.url,
        target: '_blank',
        rel: 'noopener noreferrer',
        className: styles.resultTitle,
        title: item.title,
        onClick: function(e: any) {
          e.preventDefault();
          window.open(item.url, '_blank');
        }
      }, item.title),
      metaItems.length > 0
        ? React.createElement('div', { className: styles.resultMeta }, metaItems)
        : null
    )
  );
}

function ResultsList(props: {
  items: ISpSmartItem[];
  emptyMessage: string;
}): any {
  if (props.items.length === 0) {
    return React.createElement('div', { className: styles.noResults },
      React.createElement('div', { className: styles.noResultsIcon }, '🔍'),
      React.createElement('span', null, props.emptyMessage)
    );
  }
  return React.createElement('ul', { className: styles.resultsList },
    props.items.map(function(item) {
      return React.createElement(ResultItemRow, {
        key: item.id,
        item: item
      });
    })
  );
}

export const SpSmartSearch: React.FC<ISpSmartSearchProps> = (props) => {
  var context = props.context;
  var displayMode = props.displayMode;
  var searchScope = props.searchScope;
  var scopeUrl = props.scopeUrl;
  var titleColumn = props.titleColumn;

  const [searchText, setSearchText] = useState('');
  const [includeDocuments, setIncludeDocuments] = useState(displayMode !== 'listItems');
  const [includeListItems, setIncludeListItems] = useState(displayMode !== 'documents');
  const [activeTab, setActiveTab] = useState('all' as ActiveTab);
  const [results, setResults] = useState(null as ISearchResult | null);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [hasSearched, setHasSearched] = useState(false);
  const [lastSearch, setLastSearch] = useState('');

  var effectiveDisplayMode: DisplayMode =
    includeDocuments && includeListItems ? 'both' :
    includeDocuments ? 'documents' :
    includeListItems ? 'listItems' : 'both';

  const handleSearch = useCallback(async () => {
    if (!searchText.trim()) return;
    if (!scopeUrl) {
      setErrorMessage('Please configure a Site or List URL in the web part properties.');
      return;
    }
    setIsLoading(true);
    setErrorMessage('');
    setHasSearched(true);
    setLastSearch(searchText.trim());
    try {
      var service = new SpSmartSearchService(context, titleColumn);
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
  }, [searchText, scopeUrl, searchScope, effectiveDisplayMode, context, titleColumn]);

  function handleKeyDown(e: React.KeyboardEvent<HTMLInputElement>): void {
    if (e.key === 'Enter') {
      handleSearch().catch(function(err) { console.error(err); });
    }
  }

  if (!scopeUrl) {
    return (
      <div className={styles.configRequired}>
        <h3>Configuration Required</h3>
        <p>Open the web part properties panel and configure the List or Site URL.</p>
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
          <button role="tab"
            className={styles.tab + (activeTab === 'all' ? ' ' + styles.activeTab : '')}
            onClick={function() { setActiveTab('all'); }}>
            {'All (' + allItems.length + ')'}
          </button>
          <button role="tab"
            className={styles.tab + (activeTab === 'documents' ? ' ' + styles.activeTab : '')}
            onClick={function() { setActiveTab('documents'); }}>
            {'Documents (' + docCount + ')'}
          </button>
          <button role="tab"
            className={styles.tab + (activeTab === 'listItems' ? ' ' + styles.activeTab : '')}
            onClick={function() { setActiveTab('listItems'); }}>
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
            emptyMessage: 'No results found for "' + lastSearch + '"'
          })
        ) : null}
      </div>
    </div>
  );
};

export default SpSmartSearch;
