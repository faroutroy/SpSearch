import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '../models/ISpSmartItem';

export interface ISpSmartSearchProps {
  context: WebPartContext;
  searchScope: 'site' | 'url';
  scopeUrl: string;
  titleField: string;
  displayColumn: string;        
}
