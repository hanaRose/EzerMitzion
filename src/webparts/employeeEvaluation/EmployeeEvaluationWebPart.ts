import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import EmployeeEvaluation from './components/EmployeeEvaluation';
import { IEmployeeEvaluationProps } from './components/IEmployeeEvaluationProps';

import { spfi, SPFx } from '@pnp/sp';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IEmployeeEvaluationWebPartProps {}

export default class EmployeeEvaluationWebPart
  extends BaseClientSideWebPart<IEmployeeEvaluationWebPartProps> {

  public render(): void {
    const sp = spfi().using(SPFx(this.context));

    // MSGraphClientV3 מקומי ל-render כדי שנרנדר רק אחרי קבלת הלקוח
    this.context.msGraphClientFactory.getClient('3').then((graphClient: MSGraphClientV3) => {
      const element: React.ReactElement<IEmployeeEvaluationProps> = React.createElement(EmployeeEvaluation, {
        sp,
        graphClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        context: this.context
      });

      ReactDom.render(element, this.domElement);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
