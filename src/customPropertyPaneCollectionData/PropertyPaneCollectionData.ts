import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { IPropertyPaneCollectionDataInternalProps } from './IPropertyPaneCollectionDataInternalProps';
import CollectionDataField, { ICollectionDataFieldProps } from './CollectionDataField';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPropertyPaneCollectionDataProps {
  label: string;
  value: any[];
  onPropertyChange: (propertyPath: string, newValue: any) => void;
  ctx: WebPartContext;
}

export class PropertyPaneCollectionData implements IPropertyPaneField<IPropertyPaneCollectionDataProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneCollectionDataInternalProps;
  private elem: HTMLElement;

  constructor(targetProperty: string, properties: IPropertyPaneCollectionDataProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.label,
      label: properties.label,
      value: properties.value,
      onPropertyChange: properties.onPropertyChange,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this),
      ctx: properties.ctx,
    };
  }

  public render(): void {
    if (!this.elem) {
      return;
    }

    this.onRender(this.elem);
  }

  private onDispose(element: HTMLElement): void {
    ReactDom.unmountComponentAtNode(element);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element: React.ReactElement<ICollectionDataFieldProps> = React.createElement(CollectionDataField, {
      label: this.properties.label,
      value: this.properties.value,
      onChanged: this.onChanged.bind(this),
      // required to allow the component to be re-rendered by calling this.render() externally
      stateKey: new Date().toString()
    });
    ReactDom.render(element, elem);
  }

  private onChanged(newValue: any[]): void {
    this.properties.onPropertyChange(this.targetProperty, newValue);
  }
}