 import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import Dashboard from './components/Dashboard';
import { IDashboardProps } from './components/IDashboardProps';
import { ListService } from './services/ListService';
import { HappyMomentsService } from './services/HappyMomentsService';

export default class DashboardWebPart extends BaseClientSideWebPart<{}> {
  private listService!: ListService;
  private happyMomentsService!: HappyMomentsService;

  protected async onInit(): Promise<void> {
    this.listService = new ListService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );
    this.happyMomentsService = new HappyMomentsService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );

    await Promise.all([
      this.listService.ensureList(),
      this.happyMomentsService.ensureList()
    ]);
  }

  public render(): void {
    const element = React.createElement<IDashboardProps>(Dashboard, {
      listService: this.listService,
      happyMomentsService: this.happyMomentsService
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}