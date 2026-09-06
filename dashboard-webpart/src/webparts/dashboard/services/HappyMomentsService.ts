import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const LIST_URL = '/sites/portal/Lists/HappyMoments';

export interface IHappyMomentsItem {
  Id: number;
  Title: string;
  SubTitle: string;
  EventType: string;
  Order: number;
  ExpirationDate: string | null;
}

export class HappyMomentsService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string
  ) { }

  async ensureList(): Promise<void> {
    try {
      const exists = await this.listExists();
      if (!exists) {
        await this.createList();
      }
    } catch {
      await this.createList();
    }
  }

  private async listExists(): Promise<boolean> {
    try {
      const res: SPHttpClientResponse = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/GetList('${LIST_URL}')`,
        SPHttpClient.configurations.v1
      );
      return res.ok;
    } catch {
      return false;
    }
  }

  private async createList(): Promise<void> {
    await this.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Content-Type': 'application/json;odata=nometadata' },
        body: JSON.stringify({
          Title: 'HappyMoments',
          BaseTemplate: 100,
          Description: 'Happy Moments data source'
        })
      }
    );

    const columns = [
      { FieldTypeKind: 2, Title: 'SubTitle' },
      { FieldTypeKind: 6, Title: 'EventType', Choices: { results: ['לידה', 'אירוסין', 'חתונה', 'בר/בת מצווה', 'אחר'] } },
      { FieldTypeKind: 9, Title: 'Order' }
    ];

    for (const col of columns) {
      await this.spHttpClient.post(
        `${this.siteUrl}/_api/web/GetList('${LIST_URL}')/fields`,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json;odata=nometadata' },
          body: JSON.stringify(col)
        }
      );
    }
  }

  async getItems(): Promise<IHappyMomentsItem[]> {
    const url = `${this.siteUrl}/_api/web/GetList('${LIST_URL}')/items?$select=Id,Title,SubTitle,EventType,Order0,ExpirationDate&$orderby=Order0%20asc&$top=50`;
    const res = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) return [];
    const data = await res.json();
    const now = new Date();
    const today =
      `${now.getFullYear()}-` +
      `${String(now.getMonth() + 1).padStart(2, '0')}-` +
      `${String(now.getDate()).padStart(2, '0')}`;

    return (data.value ?? [])
      .map((item: any): IHappyMomentsItem => ({
        Id: item.Id,
        Title: item.Title,
        SubTitle: item.SubTitle,
        EventType: item.EventType ?? '',
        Order: item.Order0,
        ExpirationDate: item.ExpirationDate ?? null
      }))
      .filter(
        (item: IHappyMomentsItem) =>
          !item.ExpirationDate ||
          item.ExpirationDate.substring(0, 10) >= today
      );
    }
}