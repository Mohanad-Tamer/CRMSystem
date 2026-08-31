import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import * as strings from 'SupportWebPartStrings';
import Support from './components/Support';
import type { ISupportProps } from './components/ISupportProps';

export interface ISupportWebPartProps {
  description: string;
}

export default class SupportWebPart
  extends BaseClientSideWebPart<ISupportWebPartProps> {

  public render(): void {

    const element: React.ReactElement<ISupportProps> =
      React.createElement(
        Support,
        {
          context: this.context
        }
      );

    ReactDom.render(
      element,
      this.domElement
    );
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

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
                PropertyPaneTextField(
                  'description',
                  {
                    label: strings.DescriptionFieldLabel
                  }
                )
              ]
            }
          ]
        }
      ]
    };
  }
}