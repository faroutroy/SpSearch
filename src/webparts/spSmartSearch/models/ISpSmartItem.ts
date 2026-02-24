export interface ISpSmartItem {
  id: string;
  title: string;
  type: 'document' | 'listItem';
  url: string;
  fileType?: string;
  listName?: string;
  siteUrl?: string;
  modifiedDate?: string;
  author?: string;
  summary?: string;
  businessArea?: string;
  bidDate?: string;
  estimator?: string;
  bid2WinId?: string;
  segment?: string;
  sqYards?: string;
  laneMiles?: string;
  numberOfLots?: string;
  city?: string;
  county?: string;
  state?: string;
  zipCode?: string;
  owner?: string;
  project?: string;

  // 🔥 allow dynamic property access (important)
  [key: string]: any;
}

export type DisplayMode = 'documents' | 'listItems' | 'both';
export type SearchScope = 'site' | 'url';

export interface ISearchResult {
  documents: ISpSmartItem[];
  listItems: ISpSmartItem[];
}
