import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'QuickLinksWebPartStrings';
import QuickLinks from './components/QuickLinks';
import { IQuickLinksProps } from './components/IQuickLinksProps';
import { ListService } from './services/ListService';

export interface IQuickLinksWebPartProps {
  description: string;
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
  private listService!: ListService;

  protected async onInit(): Promise<void> {
    console.log("onInit");
    this.listService = new ListService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
    await this.listService.ensureList(); // Provision list on first load
  }

  public render(): void {
    const element = React.createElement<IQuickLinksProps>(QuickLinks, {
      listService: this.listService
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}