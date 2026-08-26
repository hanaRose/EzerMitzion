import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import Cards from './components/SubjectCards';
import { ISubjectCardsProps } from './components/ISubjectCardsProps';
import { ListService } from './services/ListService';

export default class SubjectCardsWebPart extends BaseClientSideWebPart<{}> {
  private listService!: ListService;

  protected async onInit(): Promise<void> {
    console.log("onInit");
    this.listService = new ListService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl,
      this.context.pageContext.web.serverRelativeUrl
    );
    await this.listService.ensureList(); // Provision list on first load
  }

  public render(): void {
   const element = React.createElement<ISubjectCardsProps>(Cards, {
  listService: this.listService
});
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}