import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'InifiniteGridWebPartStrings';
import InifiniteGrid from './components/InifiniteGrid';
import { IInifiniteGridProps } from './components/IInifiniteGridProps';
import { sp } from "@pnp/sp";

export interface IInifiniteGridWebPartProps {
  description: string;
}

export default class InifiniteGridWebPart extends BaseClientSideWebPart<IInifiniteGridWebPartProps> {

  protected onInit(): Promise<void>{
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Infinite Grid");
    return super.onInit().then(_=> {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IInifiniteGridProps> = React.createElement(
      InifiniteGrid,
      {
        description: this.properties.description
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
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
