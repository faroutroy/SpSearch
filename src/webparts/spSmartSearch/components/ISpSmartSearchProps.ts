import { DisplayMode } from '../models/ISpSmartItem';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISpSmartSearchProps {
  context: WebPartContext;
  displayMode: DisplayMode;
  searchScope: 'site' | 'url';
  scopeUrl: string;
}
