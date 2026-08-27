import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const LIST_TITLE = 'QuickLinksList';

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
        await this.createList();
      }
    }
    catch (ex) {
      await this.createList();
    }
  }

  private async listExists(): Promise<boolean> {
    console.log("listExists");
    try {
      const res: SPHttpClientResponse = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')`,
        SPHttpClient.configurations.v1
      );
      console.log("res", res);
      return res.ok;
    }
    catch (ex) {
      //this.createList();
      return false;
    }
  }

  private async createList(): Promise<void> {
    console.log("createList");
    // Create the list
    await this.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Content-Type': 'application/json;odata=nometadata' },
        body: JSON.stringify({
          Title: LIST_TITLE,
          BaseTemplate: 100, // Generic list
          Description: 'Quick links data source'
        })
      }
    );

    // Add custom columns
    const columns = [// Single line of text
      { FieldTypeKind: 2, Title: 'Url' },
      { FieldTypeKind: 8, Title: 'IsActive' },
      { FieldTypeKind: 8, Title: 'OpenInNewTab' },
      { FieldTypeKind: 9, Title: 'Order' }
    ];

    for (const col of columns) {
      await this.spHttpClient.post(
        `${this.siteUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/fields`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json;odata=nometadata' },
          body: JSON.stringify(col)
        }
      );
    }
  }

  async getQuickLinks(): Promise<IQuickLinksItem[]> {
    console.log("getQuickLinks");
    const res: SPHttpClientResponse = await this.spHttpClient.get(
      `${this.siteUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items?$select=Id,Title,Url,Order,OpenInNewTab&$filter=IsActive eq 1&$oredrBy=Order desc`,
      SPHttpClient.configurations.v1
    );
    const data = await res.json();
    return data.value as IQuickLinksItem[];
  }
}

export interface IQuickLinksItem {
  Title: string;
  Url: string;
  Order: number;
  OpenInNewTab: boolean;
  Id: number;
}