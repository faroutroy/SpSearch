import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SpSmartSearch } from './components/SpSmartSearch';
import { ISpSmartSearchProps } from './components/ISpSmartSearchProps';
import { DisplayMode } from './models/ISpSmartItem';

export interface ISpSmartSearchWebPartProps {
  searchScope: 'site' | 'url';
  scopeUrl: string;
  titleField: string;
  displayColumn: string;
}

export default class SpSmartSearchWebPart extends BaseClientSideWebPart<ISpSmartSearchWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ISpSmartSearchProps> = React.createElement(
      SpSmartSearch,
      {
        context: this.context,
        displayMode: this.displayMode,           // kept for compatibility (edit/view mode)
        searchScope: this.properties.searchScope || 'site',
        scopeUrl: this.properties.scopeUrl || '',
        titleField: this.properties.titleField || 'Title',
        displayColumn: this.properties.displayColumn || ''
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "SharePoint Smart Search Configuration"
          },
          groups: [
            {
              groupName: "Search Settings",
              groupFields: [
                PropertyPaneTextField('titleField', {
                  label: 'Title Field',
                  description: 'Field to use as the title for search results'
                }),
                PropertyPaneDropdown('searchScope', {
                  label: 'Search Scope',
                  options: [
                    { key: 'site', text: 'Current Site' },
                    { key: 'url', text: 'Specific URL' }
                  ]
                }),
                PropertyPaneTextField('scopeUrl', {
                  label: 'Scope URL',
                  description: 'Only used when Specific URL is selected',
                  disabled: this.properties.searchScope !== 'url'
                }),
                PropertyPaneTextField('displayColumn', {
                  label: 'Column Name to Display',
                  description: 'Internal name of the column (e.g. Metro, Project_x0020_Name). This value will replace the old DisplayForm link in every result.',
                  placeholder: 'Metro'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
