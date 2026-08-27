 import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
 
const LIST_ID = '53797e3f-e921-47b0-871c-de9bd8d26396';
 
export class ListService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string
  ) { }
 
  async ensureList(): Promise<void> {
    console.log("ensureList");
    try {
      const exists = await this.listExists();
      if (!exists) {
        console.log("List not found by ID");
      }
    } catch (ex) {
      console.log("ensureList error:", ex);
    }
  }
 
  private async listExists(): Promise<boolean> {
    console.log("listExists");
    try {
      const res: SPHttpClientResponse = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/GetById('${LIST_ID}')`,
        SPHttpClient.configurations.v1
      );
      console.log("listExists res:", res.ok);
      return res.ok;
    } catch (ex) {
      return false;
    }
  }
 
  async getCards(): Promise<ICardItem[]> {
    console.log("getCards called");
    const res: SPHttpClientResponse = await this.spHttpClient.get(
      `${this.siteUrl}/_api/web/lists/GetById('${LIST_ID}')/items?$select=Id,Title,IconUrl,SubTitle,Url,IsHighlighted,IsLarge,OpenInNewTab,Order&$orderby=Order asc`,
      SPHttpClient.configurations.v1
    );
    const data = await res.json();
    console.log("getCards data:", data);
    return data.value as ICardItem[];
  }
}
 
export interface ICardItem {
  Id: number;
  Title: string;
  IconUrl: string;
  SubTitle: string;
  Url: string;
  IsHighlighted: boolean;
  IsLarge: boolean;
  OpenInNewTab: boolean;
  Order: number;
}