import { DisplayMode } from '../models/ISpSmartItem';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IColumnConfig {
  showTitle: boolean;
  showAuthor: boolean;
  showModifiedDate: boolean;
  showFileType: boolean;
  showListName: boolean;
  showBusinessArea: boolean;
  showBidDate: boolean;
  showEstimator: boolean;
  showBid2WinId: boolean;
  showSegment: boolean;
  showSqYards: boolean;
  showLaneMiles: boolean;
  showNumberOfLots: boolean;
  showCity: boolean;
  showCounty: boolean;
  showState: boolean;
  showZipCode: boolean;
  showOwner: boolean;
  showProject: boolean;
}

export interface ISpSmartSearchProps {
  context: WebPartContext;
  displayMode: DisplayMode;
  searchScope: 'site' | 'url';
  scopeUrl: string;
  titleField: string;
  columnConfig: IColumnConfig;
}
