import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

export interface INewsletterItem {
  id: number;
  title: string;
  descriptionLine1: string;
  descriptionLine2: string;
  hebrewDateLine1: string;
  hebrewDateLine2: string;
  imageUrl: string;
  linkUrl: string;
}

const INTERNAL_LIST_NAME = 'NewsletterCardsData';

export class NewsletterCardsService {
  private readonly context: WebPartContext;

  public constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getItems(): Promise<INewsletterItem[]> {
    await this.ensureList();

    const listUrl = this.getListServerRelativeUrl();

    const requestUrl =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/GetList(@listUrl)/items` +
      `?@listUrl='${encodeURIComponent(listUrl)}'` +
      `&$select=Id,Title,DescriptionLine1,DescriptionLine2,HebrewDateLine1,HebrewDateLine2,ImageUrl,LinkUrl,SortOrder,IsActive` +
      //`&$filter=IsActive eq true` +
      `&$orderby=SortOrder asc,Id asc`;

    const response = await this.context.spHttpClient.get(
      requestUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
        const errorText = await response.text();
        console.error('NewsletterCards getItems error:', errorText);

        throw new Error(
            `Failed to load newsletter items. Status: ${response.status}. Details: ${errorText}`
        );
    }

    const json = await response.json();
    console.log('NewsletterCards items from SharePoint:', json.value);

    return (json.value || []).map((item: any): INewsletterItem => ({
      id: item.Id,
      title: item.Title || '',
      descriptionLine1: item.DescriptionLine1 || '',
      descriptionLine2: item.DescriptionLine2 || '',
      hebrewDateLine1: item.HebrewDateLine1 || '',
      hebrewDateLine2: item.HebrewDateLine2 || '',
      imageUrl: item.ImageUrl?.Url || '',
      linkUrl: item.LinkUrl?.Url || '#'
    }));
  }

  private async ensureList(): Promise<void> {
    const exists = await this.doesListExist();

    if (!exists) {
      await this.createList();
    }

    await this.ensureFields();
  }

  private async doesListExist(): Promise<boolean> {
    const listUrl = this.getListServerRelativeUrl();

    const requestUrl =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/GetList(@listUrl)` +
      `?@listUrl='${encodeURIComponent(listUrl)}'`;

    const response = await this.context.spHttpClient.get(
      requestUrl,
      SPHttpClient.configurations.v1
    );

    return response.ok;
  }

  private async createList(): Promise<void> {
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;

    const body = {
      BaseTemplate: 100,
      Title: INTERNAL_LIST_NAME,
      Description: 'Newsletter cards data list',
      AllowContentTypes: true
    };

    const response = await this.context.spHttpClient.post(
      requestUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(body)
      }
    );

    if (!response.ok) {
        const errorText = await response.text();
        console.error('NewsletterCards createList error:', errorText);

        throw new Error(
            `Failed to create list. Status: ${response.status}. Details: ${errorText}`
        );
    }
  }

  private async ensureFields(): Promise<void> {
    await this.ensureField(
      'DescriptionLine1',
      '<Field Type="Note" DisplayName="DescriptionLine1" Name="DescriptionLine1" StaticName="DescriptionLine1" NumLines="3" RichText="FALSE" />'
    );

    await this.ensureField(
      'DescriptionLine2',
      '<Field Type="Note" DisplayName="DescriptionLine2" Name="DescriptionLine2" StaticName="DescriptionLine2" NumLines="3" RichText="FALSE" />'
    );

    await this.ensureField(
      'HebrewDateLine1',
      '<Field Type="Text" DisplayName="HebrewDateLine1" Name="HebrewDateLine1" StaticName="HebrewDateLine1" MaxLength="255" />'
    );

    await this.ensureField(
      'HebrewDateLine2',
      '<Field Type="Text" DisplayName="HebrewDateLine2" Name="HebrewDateLine2" StaticName="HebrewDateLine2" MaxLength="255" />'
    );

    await this.ensureField(
      'ImageUrl',
      '<Field Type="URL" DisplayName="Image Url" Name="ImageUrl" StaticName="ImageUrl" Format="Hyperlink" />'
    );

    await this.ensureField(
      'LinkUrl',
      '<Field Type="URL" DisplayName="Link Url" Name="LinkUrl" StaticName="LinkUrl" Format="Hyperlink" />'
    );

    await this.ensureField(
      'SortOrder',
      '<Field Type="Number" DisplayName="SortOrder" Name="SortOrder" StaticName="SortOrder" Min="0" Decimals="0" />'
    );

    await this.ensureField(
      'IsActive',
      '<Field Type="Boolean" DisplayName="IsActive" Name="IsActive" StaticName="IsActive"><Default>1</Default></Field>'
    );
  }

  private async ensureField(internalName: string, schemaXml: string): Promise<void> {
  const exists = await this.doesFieldExist(internalName);

  if (exists) {
    return;
  }

  const requestUrl =
    `${this.context.pageContext.web.absoluteUrl}` +
    `/_api/web/GetList(@listUrl)/fields/CreateFieldAsXml` +
    `?@listUrl='${this.getListUrlParameter()}'`;

  const body = {
    parameters: {
        SchemaXml: schemaXml,
        Options: 8
    }
};

  const response = await this.context.spHttpClient.post(
    requestUrl,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      body: JSON.stringify(body)
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    console.error(`NewsletterCards ensureField error for ${internalName}:`, errorText);

    throw new Error(
      `Failed to create field ${internalName}. Status: ${response.status}. Details: ${errorText}`
    );
  }
}

  private async doesFieldExist(internalName: string): Promise<boolean> {
    const listUrl = this.getListServerRelativeUrl();

    const requestUrl =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/GetList(@listUrl)/fields/getbyinternalnameortitle('${internalName}')` +
      `?@listUrl='${encodeURIComponent(listUrl)}'`;

    const response = await this.context.spHttpClient.get(
      requestUrl,
      SPHttpClient.configurations.v1
    );

    return response.ok;
  }

  private getListServerRelativeUrl(): string {
  const webServerRelativeUrl = this.context.pageContext.web.serverRelativeUrl;
  const cleanWebUrl = webServerRelativeUrl === '/' ? '' : webServerRelativeUrl;

  return `${cleanWebUrl}/Lists/${INTERNAL_LIST_NAME}`;
}

private getListUrlParameter(): string {
  return this.getListServerRelativeUrl().replace(/'/g, "''");
}
}