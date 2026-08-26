 import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import styles from "./SiteHero.module.scss";
import { ISiteHeroProps } from "./ISiteHeroProps";

interface ISiteHeroItem {
  Id: number;
  Title: string;
  HeroText: string;
  SubTitle: string;
  BackgroundImageUrl: string;
  LogoImageUrl: string;
  ButtonText: string;
  ButtonUrl: string;
}

const SiteHero: React.FC<ISiteHeroProps> = ({ context, listName, overrideHeroText, overrideSubTitle }) => {
  const [item, setItem] = useState<ISiteHeroItem | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>("");

  const webUrl = context.pageContext.web.absoluteUrl;

  const getSafeListName = (name: string): string => name.replace(/'/g, "''");

  const getJsonHeaders = (): HeadersInit => ({
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    "odata-version": ""
  });

  const listExists = async (name: string): Promise<boolean> => {
    const safeListName = getSafeListName(name);
    const url = `${webUrl}/_api/web/lists/getbytitle('${safeListName}')?$select=Id`;
    const response: SPHttpClientResponse = await context.spHttpClient.get(
      url, SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata.metadata=none" } }
    );
    if (response.ok) return true;
    if (response.status === 404) return false;
    throw new Error(`List check failed: ${response.status}`);
  };

  const createList = async (name: string): Promise<void> => {
    const trimmedName = (name || "").trim();
    if (!trimmedName) throw new Error("List name is empty.");
    const url = `${webUrl}/_api/web/lists`;
    const body = {
      __metadata: { type: "SP.List" },
      AllowContentTypes: true,
      BaseTemplate: 100,
      ContentTypesEnabled: true,
      Description: "Header content for the SiteHero web part",
      Title: trimmedName
    };
    const response: SPHttpClientResponse = await context.spHttpClient.post(
      url, SPHttpClient.configurations.v1,
      { headers: getJsonHeaders(), body: JSON.stringify(body) }
    );
    const responseText = await response.text();
    if (!response.ok) throw new Error(`List creation failed: ${response.status} | ${responseText}`);
  };

  const fieldExists = async (name: string, internalName: string): Promise<boolean> => {
    const safeListName = getSafeListName((name || "").trim());
    const safeInternalName = (internalName || "").trim().replace(/'/g, "''");
    const url =
      `${webUrl}/_api/web/lists/getbytitle('${safeListName}')/fields` +
      `?$select=Id,InternalName,Title` +
      `&$filter=(InternalName eq '${safeInternalName}' or Title eq '${safeInternalName}')` +
      `&$top=1`;
    const response: SPHttpClientResponse = await context.spHttpClient.get(
      url, SPHttpClient.configurations.v1, { headers: getJsonHeaders() }
    );
    const responseText = await response.text();
    if (!response.ok) throw new Error(`Field check failed for ${internalName}: ${response.status} | ${responseText}`);
    const data: any = responseText ? JSON.parse(responseText) : null;
    const results = data?.d?.results ?? data?.value ?? [];
    return Array.isArray(results) && results.length > 0;
  };

  const createFieldFromXml = async (name: string, schemaXml: string): Promise<void> => {
    const safeListName = getSafeListName(name);
    const url = `${webUrl}/_api/web/lists/getbytitle('${safeListName}')/fields/createfieldasxml`;
    const body = {
      parameters: {
        __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
        Options: 0,
        SchemaXml: schemaXml
      }
    };
    const response: SPHttpClientResponse = await context.spHttpClient.post(
      url, SPHttpClient.configurations.v1,
      { headers: getJsonHeaders(), body: JSON.stringify(body) }
    );
    if (!response.ok) throw new Error(`Field creation failed: ${response.status}`);
  };

  const ensureField = async (name: string, internalName: string, schemaXml: string): Promise<void> => {
    const exists = await fieldExists(name, internalName);
    if (!exists) await createFieldFromXml(name, schemaXml);
  };

  const ensureSiteHeaderList = async (name: string): Promise<void> => {
    const exists = await listExists(name);
    if (!exists) await createList(name);

    await ensureField(name, "HeroText",
      `<Field Type="Note" DisplayName="HeroText" Name="HeroText" StaticName="HeroText" RichText="FALSE" NumLines="6" Group="SiteHero Columns" />`);
    await ensureField(name, "SubTitle",
      `<Field Type="Text" DisplayName="SubTitle" Name="SubTitle" StaticName="SubTitle" Group="SiteHero Columns" />`);
    await ensureField(name, "BackgroundImageUrl",
      `<Field Type="Text" DisplayName="BackgroundImageUrl" Name="BackgroundImageUrl" StaticName="BackgroundImageUrl" Group="SiteHero Columns" />`);
    await ensureField(name, "LogoImageUrl",
      `<Field Type="Text" DisplayName="LogoImageUrl" Name="LogoImageUrl" StaticName="LogoImageUrl" Group="SiteHero Columns" />`);
    await ensureField(name, "ButtonText",
      `<Field Type="Text" DisplayName="ButtonText" Name="ButtonText" StaticName="ButtonText" Group="SiteHero Columns" />`);
    await ensureField(name, "ButtonUrl",
      `<Field Type="Text" DisplayName="ButtonUrl" Name="ButtonUrl" StaticName="ButtonUrl" Group="SiteHero Columns" />`);
  };

  useEffect(() => {
    const loadItem = async (): Promise<void> => {
      try {
        setLoading(true);
        setError("");
        const effectiveListName = (listName || "SiteHeader").trim();
        await ensureSiteHeaderList(effectiveListName);
        const safeListName = getSafeListName(effectiveListName);
        const url =
          `${webUrl}/_api/web/lists/getbytitle('${safeListName}')/items` +
          `?$select=Id,Title,HeroText,SubTitle,BackgroundImageUrl,LogoImageUrl,ButtonText,ButtonUrl` +
          `&$top=1`;
        const response: SPHttpClientResponse = await context.spHttpClient.get(
          url, SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata.metadata=none" } }
        );
        if (!response.ok) throw new Error(`Request failed: ${response.status}`);
        const data = await response.json();
        const firstItem = data.value && data.value.length > 0 ? data.value[0] : null;
        setItem(firstItem);
      } catch (err) {
        const message = err instanceof Error ? err.message : "Failed to load header data.";
        setError(message);
      } finally {
        setLoading(false);
      }
    };
    loadItem().catch(() => {
      setError("Failed to load header data.");
      setLoading(false);
    });
  }, [context, listName]);

  if (loading) return null;
  if (error) return <div className={styles.message}>{error}</div>;
  if (!item) return <div className={styles.message}>No active header item found.</div>;

  const backgroundStyle: React.CSSProperties = {
    backgroundImage: `url('${item.BackgroundImageUrl}')`
  };

  const heroText = overrideHeroText || item.HeroText || item.Title;
  const subTitle = overrideSubTitle || item.SubTitle;

  return (
    <section className={styles.header} style={backgroundStyle} dir="rtl">
      <div className={styles.container}>
        <div className={styles.textGroup}>
          {item.LogoImageUrl ? (
            <img className={styles.logo} src={item.LogoImageUrl} alt={item.Title || "Logo"} />
          ) : null}

          <h1 className={styles.title}>{heroText}</h1>

          {subTitle ? (
            <p className={styles.subtitle}>{subTitle}</p>
          ) : null}
        </div>

        {item.ButtonText && item.ButtonUrl ? (
          <a className={styles.button} href={item.ButtonUrl} target="_blank" rel="noreferrer">
            <span>{item.ButtonText}</span>
          </a>
        ) : null}
      </div>
    </section>
  );
};

export default SiteHero;