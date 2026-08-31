import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import Execution from './components/Execution';
import { IExecutionProps } from './components/IExecutionProps';

export interface IExecutionWebPartProps {
  description: string;
}

export default class ExecutionWebPart
  extends BaseClientSideWebPart<IExecutionWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IExecutionProps> =
      React.createElement(
        Execution,
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
    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(
      this.domElement
    );
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}