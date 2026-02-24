import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '../models/ISpSmartItem';

export interface ISpSmartSearchProps {
  context: WebPartContext;
  displayMode: DisplayMode;
  searchScope: 'site' | 'url';
  scopeUrl: string;
  titleField: string;
  displayColumn: string;        // ← this is the only setting now
}
