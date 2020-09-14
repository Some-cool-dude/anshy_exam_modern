import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AnshyWebPartWebPartStrings';
import AnshyWebPart from './components/AnshyWebPart';
import IAnshyWebPartProps from './components/AnshyWebPart';

export interface IAnshyWebPartWebPartProps {
  description: string;
}

export default class AnshyWebPartWebPart extends BaseClientSideWebPart<IAnshyWebPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAnshyWebPartProps> = React.createElement(
      AnshyWebPart,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
