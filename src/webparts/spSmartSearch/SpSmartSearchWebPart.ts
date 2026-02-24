import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SpSmartSearch } from './components/SpSmartSearch';
import { ISpSmartSearchProps } from './components/ISpSmartSearchProps';
import { DisplayMode } from './models/ISpSmartItem';

export interface ISpSmartSearchWebPartProps {
  displayMode: DisplayMode;
  searchScope: 'site' | 'url';
  scopeUrl: string;
  titleColumn: string;
  displayColumns: string;
}

export default class SpSmartSearchWebPart extends BaseClientSideWebPart<ISpSmartSearchWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ISpSmartSearchProps> = React.createElement(
      SpSmartSearch,
      {
        context: this.context,
        displayMode: this.properties.displayMode || 'listItems',
        searchScope: this.properties.searchScope || 'url',
        scopeUrl: this.properties.scopeUrl || '',
        titleColumn: this.properties.titleColumn || '',
        displayColumns: this.properties.displayColumns || ''
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const displayModeOptions: IPropertyPaneDropdownOption[] = [
      { key: 'both', text: 'Both (Documents & List Items with tabs)' },
      { key: 'documents', text: 'Documents only' },
      { key: 'listItems', text: 'List Items only' }
    ];

    const searchScopeOptions: IPropertyPaneDropdownOption[] = [
      { key: 'site', text: 'Entire Site (all lists & libraries)' },
      { key: 'url', text: 'Specific List or Library URL' }
    ];

    const isSiteScope = !this.properties.searchScope || this.properties.searchScope === 'site';

    return {
      pages: [
        {
          header: { description: 'Configure search settings and display options.' },
          groups: [
            {
              groupName: '📋 Display Options',
              groupFields: [
                PropertyPaneDropdown('displayMode', {
                  label: 'Content types to display',
                  options: displayModeOptions,
                  selectedKey: this.properties.displayMode || 'listItems'
                })
              ]
            },
            {
              groupName: '🔍 Search Scope',
              groupFields: [
                PropertyPaneDropdown('searchScope', {
                  label: 'Search scope',
                  options: searchScopeOptions,
                  selectedKey: this.properties.searchScope || 'url'
                }),
                PropertyPaneTextField('scopeUrl', {
                  label: isSiteScope ? 'Site URL' : 'List or Library URL',
                  description: isSiteScope
                    ? 'Enter the full URL of the SharePoint site to search'
                    : 'Enter the full URL of a specific list or document library',
                  placeholder: isSiteScope
                    ? 'https://contoso.sharepoint.com/sites/MySite'
                    : 'https://contoso.sharepoint.com/sites/MySite/Lists/MyList',
                  multiline: false,
                  resizable: false
                })
              ]
            },
            {
              groupName: '🏷 Column Settings',
              groupFields: [
                PropertyPaneTextField('titleColumn', {
                  label: 'Title Column',
                  description: 'Internal column name to show as the result title. Example: Bid2WinID, Project, Title',
                  placeholder: 'e.g. Bid2WinID'
                }),
                PropertyPaneLabel('titleColumnNote', {
                  text: '💡 The value of this column will appear as the title for each result row.'
                }),
                PropertyPaneTextField('displayColumns', {
                  label: 'Additional Columns to Display',
                  description: 'Enter internal column names separated by commas to show as metadata under each result.',
                  placeholder: 'e.g. Estimator,City,State,BidDate,Segment',
                  multiline: true,
                  resizable: false,
                  rows: 4
                }),
                PropertyPaneLabel('displayColumnsNote', {
                  text: '💡 Use exact internal column names from your SharePoint list. Separate multiple columns with commas.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
