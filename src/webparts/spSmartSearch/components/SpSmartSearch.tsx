import * as React from 'react';
import { useState, useCallback } from 'react';
import styles from './SpSmartSearch.module.scss';
import { ISpSmartSearchProps } from './ISpSmartSearchProps';
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
  try { return new Date(dateStr).toLocaleDateString(); }
  catch { return ''; }
}

const ResultItemRow: React.FC<{ item: ISpSmartItem }> = ({ item }) => (
  <li className={styles.resultItem}>
    <div className={styles.resultIcon}>
      {item.type === 'listItem' ? '📋' : getFileIcon(item.fileType)}
    </div>
    <div className={styles.resultContent}>
      <a href={item.url} target="_blank" rel="noopener noreferrer"
        className={styles.resultTitle} title={item.title}>
        {item.title}
      </a>
      <div className={styles.resultMeta}>
        {item.fileType && <span>{item.fileType.toUpperCase()}</span>}
        {item.listName && <span>{item.listName}</span>}
        {item.author && <span>By {item.author}</span>}
        {item.modifiedDate && <span>Modified {formatDate(item.modifiedDate)}</span>}
      </div>
      {item.summary && (
        <div className={styles.resultSummary}
          dangerouslySetInnerHTML={{ __html: item.summary }} />
      )}
    </div>
  </li>
);

const ResultsList: React.FC<{ items: ISpSmartItem[]; emptyMessage: string }> = ({ items, emptyMessage }) => {
  if (items.length === 0) {
    return (
      <div className={styles.noResults}>
        <div className={styles.noResultsIcon}>🔍</div>
        <span>{emptyMessage}</span>
      </div>
    );
  }
  return (
    <ul className={styles.resultsList}>
      {items.map(item => <ResultItemRow key={item.id} item={item} />)}
    </ul>
  );
};

export const SpSmartSearch: React.FC<ISpSmartSearchProps> = (props) => {
  const { context, displayMode, searchScope, scopeUrl } = props;

  const [searchText, setSearchText] = useState<string>('');
  const [includeDocuments, setIncludeDocuments] = useState<boolean>(displayMode !== 'listItems');
  const [includeListItems, setIncludeListItems] = useState<boolean>(displayMode !== 'documents');
  const [activeTab, setActiveTab] = useState<ActiveTab>('all');
  const [results, setResults] = useState<ISearchResult | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [hasSearched, setHasSearched] = useState<boolean>(false);

  const effectiveDisplayMode: DisplayMode =
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
      const service = new SpSmartSearchService(context);
      const searchResults = await service.search(searchText.trim(), effectiveDisplayMode, searchScope, scopeUrl);
      setResults(searchResults);
      if (effectiveDisplayMode === 'both') setActiveTab('all');
      else if (effectiveDisplayMode === 'documents') setActiveTab('documents');
      else setActiveTab('listItems');
    } catch (err: any) {
      setErrorMessage(err?.message || 'An unexpected error occurred while searching.');
      setResults(null);
    } finally {
      setIsLoading(false);
    }
  }, [searchText, scopeUrl, searchScope, effectiveDisplayMode, context]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') { handleSearch().catch(console.error); }
  };

  if (!scopeUrl) {
    return (
      <div className={styles.configRequired}>
        <h3>Configuration Required</h3>
        <p>Open the web part properties panel and provide a Site URL or List/Library URL to enable search.</p>
      </div>
    );
  }

  const allItems: ISpSmartItem[] = results
    ? [...(results.documents || []), ...(results.listItems || [])]
    : [];

  const visibleItems: ISpSmartItem[] =
    activeTab === 'documents' ? (results?.documents || []) :
    activeTab === 'listItems' ? (results?.listItems || []) :
    allItems;

  const showTabs = effectiveDisplayMode === 'both';

  return (
    <div className={styles.spSmartSearch}>
      <div className={styles.searchBar}>
        <input
          type="text"
          className={styles.searchInput}
          placeholder="Enter search terms…"
          value={searchText}
          onChange={e => setSearchText(e.target.value)}
          onKeyDown={handleKeyDown}
          aria-label="Search query"
        />
        <div className={styles.checkboxGroup}>
          {displayMode !== 'listItems' && (
            <label>
              <input type="checkbox" checked={includeDocuments}
                onChange={e => setIncludeDocuments(e.target.checked)} />
              Documents
            </label>
          )}
          {displayMode !== 'documents' && (
            <label>
              <input type="checkbox" checked={includeListItems}
                onChange={e => setIncludeListItems(e.target.checked)} />
              List Items
            </label>
          )}
          {displayMode === 'both' && (
            <label>
              <input type="checkbox" checked={includeDocuments && includeListItems}
                onChange={e => { setIncludeDocuments(e.target.checked); setIncludeListItems(e.target.checked); }} />
              Both
            </label>
          )}
        </div>
        <button className={styles.searchButton}
          onClick={() => { handleSearch().catch(console.error); }}
          disabled={isLoading || !searchText.trim()}
          aria-label="Search">
          {isLoading ? 'Searching…' : 'Search'}
        </button>
      </div>

      {errorMessage && <div className={styles.errorMessage} role="alert">{errorMessage}</div>}

      {showTabs && results && (
        <div className={styles.tabContainer} role="tablist">
          <button role="tab" aria-selected={activeTab === 'all'}
            className={`${styles.tab} ${activeTab === 'all' ? styles.activeTab : ''}`}
            onClick={() => setActiveTab('all')}>
            All ({allItems.length})
          </button>
          <button role="tab" aria-selected={activeTab === 'documents'}
            className={`${styles.tab} ${activeTab === 'documents' ? styles.activeTab : ''}`}
            onClick={() => setActiveTab('documents')}>
            Documents ({results.documents?.length || 0})
          </button>
          <button role="tab" aria-selected={activeTab === 'listItems'}
            className={`${styles.tab} ${activeTab === 'listItems' ? styles.activeTab : ''}`}
            onClick={() => setActiveTab('listItems')}>
            List Items ({results.listItems?.length || 0})
          </button>
        </div>
      )}

      <div className={styles.resultsContainer}>
        {isLoading && <div className={styles.loadingSpinner} aria-live="polite">Searching…</div>}
        {!isLoading && !hasSearched && (
          <div className={styles.emptyState}>Enter a search term and press Search to get started.</div>
        )}
        {!isLoading && hasSearched && results && (
          <ResultsList items={visibleItems}
            emptyMessage={`No ${activeTab === 'documents' ? 'documents' : activeTab === 'listItems' ? 'list items' : 'results'} found for "${searchText}"`} />
        )}
      </div>
    </div>
  );
};

export default SpSmartSearch;
