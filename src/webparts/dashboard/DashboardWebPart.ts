import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import * as strings from 'DashboardWebPartStrings';

import Dashboard from './components/Dashboard';
import {
  IDashboardProps
} from './components/IDashboardProps';

export interface IDashboardWebPartProps {
  description: string;
}

export default class DashboardWebPart
  extends BaseClientSideWebPart<IDashboardWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IDashboardProps> =
      React.createElement(
        Dashboard,
        {
          context: this.context
        }
      );

    ReactDom.render(
      element,
      this.domElement
    );
  }

  protected onDispose(): void {

    ReactDom.unmountComponentAtNode(
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
            description:
              strings.PropertyPaneDescription
          },

          groups: [

            {

              groupName:
                strings.BasicGroupName,

              groupFields: [

                PropertyPaneTextField(
                  'description',
                  {
                    label:
                      strings.DescriptionFieldLabel
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