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

import Opportunities from './components/Opportunities';
import type { IOpportunitiesProps } from './components/IOpportunitiesProps';

export interface IOpportunitiesWebPartProps {
  description: string;
}

export default class OpportunitiesWebPart
  extends BaseClientSideWebPart<IOpportunitiesWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IOpportunitiesProps> =
      React.createElement(
        Opportunities,
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
            description: 'Opportunities'
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