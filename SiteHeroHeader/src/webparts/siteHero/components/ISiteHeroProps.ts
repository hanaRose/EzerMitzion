 import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISiteHeroProps {
  context: WebPartContext;
  listName: string;
  overrideHeroText?: string;
  overrideSubTitle?: string;
}