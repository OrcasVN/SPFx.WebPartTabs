import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WebPartTabsWebPartStrings';
import WebPartTabs from './components/WebPartTabs';
import { collectionTab, IWebPartTabsProps, tabStyle } from './components/IWebPartTabsProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import WebPartTabsServices from './services';
export interface IWebPartTabsWebPartProps {
  collectionTabs: collectionTab[];
  tabStyle: tabStyle;
}

export default class WebPartTabsWebPart extends BaseClientSideWebPart<IWebPartTabsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private options: any[] = [];

  public render(): void {
    const element: React.ReactElement<IWebPartTabsProps> = React.createElement(
      WebPartTabs,
      {
        tabStyle: this.properties.tabStyle,
        collectionTabs: this.properties.collectionTabs,
        wpContext: this.context,
        displayMode: this.displayMode,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    this.getWebParts().then(() => { }).catch(e => console.log(e))
    return super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async getWebParts() {
    const service = new WebPartTabsServices(this.context as WebPartContext)
    this.options = await service.getWebParts()
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('tabStyle', {
                  label: 'Tab Style',
                  options: [
                    { key: 'links', text: 'Links' },
                    { key: 'tabs', text: 'Tabs' }
                  ],
                  selectedKey: this.properties.tabStyle,
                }),
                PropertyFieldCollectionData("collectionTabs", {
                  key: "collectionTabs",
                  label: "Collection Tabs",
                  panelHeader: "Collection tabs panel header",
                  manageBtnLabel: "Manage collection tabs",
                  value: this.properties.collectionTabs,
                  fields: [
                    {
                      id: "Title",
                      title: "Title",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "WebPart",
                      title: "Web Part",
                      type: CustomCollectionFieldType.dropdown,
                      options: this.options,
                      required: true
                    },
                    {
                      id: "DisplayOrder",
                      title: "Display Order",
                      type: CustomCollectionFieldType.number
                    },
                    {
                      id: "IconName",
                      title: "Icon Name",
                      type: CustomCollectionFieldType.string,
                    },
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
