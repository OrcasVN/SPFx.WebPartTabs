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
import { PropertyPaneCollectionData } from '../../customPropertyPaneCollectionData/PropertyPaneCollectionData';
import { update } from '@microsoft/sp-lodash-subset';
export interface IWebPartTabsWebPartProps {
  collectionTabs: collectionTab[];
  tabStyle: tabStyle;
  fontSize: string;
}

export default class WebPartTabsWebPart extends BaseClientSideWebPart<IWebPartTabsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IWebPartTabsProps> = React.createElement(
      WebPartTabs,
      {
        tabStyle: this.properties.tabStyle,
        collectionTabs: this.properties.collectionTabs,
        wpContext: this.context,
        displayMode: this.displayMode,
        fontSize: this.properties.fontSize,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
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


  private onValueChanged(propertyPath: string, newValue: any) {
    update(this.properties, propertyPath, (): any => { return newValue; });
    this.render();
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
                new PropertyPaneCollectionData('collectionTabs', {
                  label: 'Collection Tabs',
                  ctx: this.context,
                  onPropertyChange: this.onValueChanged.bind(this),
                  value: this.properties.collectionTabs,
                }),
                PropertyPaneTextField('fontSize', {
                  label: 'Font size',
                  value: this.properties.fontSize,
                  placeholder: 'Example: 16, 16px, 16em, 16rem...',
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
