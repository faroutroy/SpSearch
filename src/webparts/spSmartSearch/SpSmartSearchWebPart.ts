import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneToggle,
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
  // Column visibility toggles
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

export default class SpSmartSearchWebPart extends BaseClientSideWebPart<ISpSmartSearchWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ISpSmartSearchProps> = React.createElement(
      SpSmartSearch,
      {
        context: this.context,
        displayMode: this.properties.displayMode || 'both',
        searchScope: this.properties.searchScope || 'site',
        scopeUrl: this.properties.scopeUrl || '',
        // Pass column visibility to component
        columnConfig: {
          showTitle: this.properties.showTitle !== false,
          showAuthor: this.properties.showAuthor !== false,
          showModifiedDate: this.properties.showModifiedDate !== false,
          showFileType: this.properties.showFileType !== false,
          showListName: this.properties.showListName !== false,
          showBusinessArea: this.properties.showBusinessArea !== false,
          showBidDate: this.properties.showBidDate !== false,
          showEstimator: this.properties.showEstimator !== false,
          showBid2WinId: this.properties.showBid2WinId !== false,
          showSegment: this.properties.showSegment !== false,
          showSqYards: this.properties.showSqYards !== false,
          showLaneMiles: this.properties.showLaneMiles !== false,
          showNumberOfLots: this.properties.showNumberOfLots !== false,
          showCity: this.properties.showCity !== false,
          showCounty: this.properties.showCounty !== false,
          showState: this.properties.showState !== false,
          showZipCode: this.properties.showZipCode !== false,
          showOwner: this.properties.showOwner !== false,
          showProject: this.properties.showProject !== false,
        }
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
                  selectedKey: this.properties.displayMode || 'both'
                }),
                PropertyPaneLabel('displayModeDescription', {
                  text: this.getDisplayModeDescription()
                })
              ]
            },
            {
              groupName: '🔍 Search Scope',
              groupFields: [
                PropertyPaneDropdown('searchScope', {
                  label: 'Search scope',
                  options: searchScopeOptions,
                  selectedKey: this.properties.searchScope || 'site'
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
                }),
                PropertyPaneLabel('scopeNote', {
                  text: isSiteScope
                    ? '💡 Searches all lists and libraries within this site.'
                    : '💡 Searches only the list or library at the URL above.'
                })
              ]
            },
            {
              groupName: '🗂 Columns to Display',
              groupFields: [
                PropertyPaneLabel('columnsNote', {
                  text: 'Toggle which columns appear in search results.'
                }),
                PropertyPaneToggle('showTitle', {
                  label: 'Title',
                  checked: this.properties.showTitle !== false
                }),
                PropertyPaneToggle('showAuthor', {
                  label: 'Author',
                  checked: this.properties.showAuthor !== false
                }),
                PropertyPaneToggle('showModifiedDate', {
                  label: 'Modified Date',
                  checked: this.properties.showModifiedDate !== false
                }),
                PropertyPaneToggle('showFileType', {
                  label: 'File Type',
                  checked: this.properties.showFileType !== false
                }),
                PropertyPaneToggle('showListName', {
                  label: 'List / Site Name',
                  checked: this.properties.showListName !== false
                }),
                PropertyPaneToggle('showBusinessArea', {
                  label: 'Business Area',
                  checked: this.properties.showBusinessArea !== false
                }),
                PropertyPaneToggle('showBidDate', {
                  label: 'Bid Date',
                  checked: this.properties.showBidDate !== false
                }),
                PropertyPaneToggle('showEstimator', {
                  label: 'Estimator',
                  checked: this.properties.showEstimator !== false
                }),
                PropertyPaneToggle('showBid2WinId', {
                  label: 'Bid2Win ID',
                  checked: this.properties.showBid2WinId !== false
                }),
                PropertyPaneToggle('showSegment', {
                  label: 'Segment',
                  checked: this.properties.showSegment !== false
                }),
                PropertyPaneToggle('showSqYards', {
                  label: 'Sq Yards',
                  checked: this.properties.showSqYards !== false
                }),
                PropertyPaneToggle('showLaneMiles', {
                  label: 'Lane Miles',
                  checked: this.properties.showLaneMiles !== false
                }),
                PropertyPaneToggle('showNumberOfLots', {
                  label: 'Number of Lots',
                  checked: this.properties.showNumberOfLots !== false
                }),
                PropertyPaneToggle('showCity', {
                  label: 'City',
                  checked: this.properties.showCity !== false
                }),
                PropertyPaneToggle('showCounty', {
                  label: 'County',
                  checked: this.properties.showCounty !== false
                }),
                PropertyPaneToggle('showState', {
                  label: 'State',
                  checked: this.properties.showState !== false
                }),
                PropertyPaneToggle('showZipCode', {
                  label: 'Zip Code',
                  checked: this.properties.showZipCode !== false
                }),
                PropertyPaneToggle('showOwner', {
                  label: 'Owner',
                  checked: this.properties.showOwner !== false
                }),
                PropertyPaneToggle('showProject', {
                  label: 'Project',
                  checked: this.properties.showProject !== false
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private getDisplayModeDescription(): string {
    switch (this.properties.displayMode) {
      case 'documents': return '📄 Only documents will be searched and shown.';
      case 'listItems': return '📋 Only list items will be searched and shown.';
      default: return '🗂 Both shown in three tabs: Documents | List Items | All.';
    }
  }
}
