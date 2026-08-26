 import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const LIST_URL = '/sites/portal/Lists/EmployeeDirectory';

export interface IEmployee {
  Id: number;
  Title: string;
  LastName: string;
  Extension: string;
  Email: string;
  Department: string;
  Branch: string;
}

export class ListService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string,
    private serverRelativeUrl: string
  ) { }

  async ensureList(): Promise<void> {
    try {
      const exists = await this.listExists();
      if (!exists) await this.createList();
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
          Title: 'EmployeeDirectory',
          BaseTemplate: 100,
          Description: 'Employee Directory'
        })
      }
    );

    const columns = [
      { FieldTypeKind: 2, Title: 'LastName' },
      { FieldTypeKind: 2, Title: 'Extension' },
      { FieldTypeKind: 2, Title: 'Email' },
      { FieldTypeKind: 2, Title: 'Department' },
      { FieldTypeKind: 2, Title: 'Branch' },
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

  async getEmployees(): Promise<IEmployee[]> {
    const res: SPHttpClientResponse = await this.spHttpClient.get(
      `${this.siteUrl}/_api/web/GetList('${LIST_URL}')/items` +
      `?$select=Id,Title,LastName,Extension,Email,Department,Branch&$orderby=LastName asc&$top=5000`,
      SPHttpClient.configurations.v1
    );
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value ?? []) as IEmployee[];
  }
}