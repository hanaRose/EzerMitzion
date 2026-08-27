import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const LIST_TITLE = 'AdditionalInfoButtonList';

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
          Description: 'Additional Info Button data source'
        })
      }
    );

    // Add custom columns
    const columns = [// Single line of text
      { FieldTypeKind: 2, Title: 'Url' }
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

  async getAdditionalInfoButtonData(): Promise<IAdditionalInfoButtonItem[]> {
    console.log("getQuickLinks");
    const res: SPHttpClientResponse = await this.spHttpClient.get(
      `${this.siteUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items?$select=Id,Title,Url&$top=1`,
      SPHttpClient.configurations.v1
    );
    const data = await res.json();
    return data.value as IAdditionalInfoButtonItem[];
  }
}

export interface IAdditionalInfoButtonItem {
  Title: string;
  Url: string;
  Id: number;
}