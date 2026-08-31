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

import Tendering from './components/Tendering';
import type { ITenderingProps } from './components/ITenderingProps';

export interface ITenderingWebPartProps {
  description: string;
}

export default class TenderingWebPart
  extends BaseClientSideWebPart<ITenderingWebPartProps> {

  public render(): void {

    const element: React.ReactElement<ITenderingProps> =
      React.createElement(
        Tendering,
        {
          context: this.context
        }
      );

    ReactDom.render(
      element,
      this.domElement
    );
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
            description: 'Tendering'
          },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneTextField(
                  'description',
                  {
                    label: 'Description'
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