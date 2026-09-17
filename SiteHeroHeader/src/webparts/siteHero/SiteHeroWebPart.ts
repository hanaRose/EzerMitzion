import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import {
  PropertyPaneButton,
  PropertyPaneButtonType
} from "@microsoft/sp-property-pane";
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";

import SiteHero from "./components/SiteHero";
import { ISiteHeroProps } from "./components/ISiteHeroProps";

export interface ISiteHeroWebPartProps {
  heroText: string;
  subTitle: string;
  backgroundImageUrl: string;
}

export default class SiteHeroWebPart
  extends BaseClientSideWebPart<ISiteHeroWebPartProps> {

  public render(): void {
    const mainContent = document.querySelector(
  "section.mainContent"
) as HTMLElement | null;

if (mainContent) {
  mainContent.style.marginTop = "-26px";
}
    const element: React.ReactElement<ISiteHeroProps> =
      React.createElement(SiteHero, {
        heroText: this.properties.heroText,
        subTitle: this.properties.subTitle,
        backgroundImageUrl: this.properties.backgroundImageUrl
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
  const mainContent = document.querySelector(
    "section.mainContent"
  ) as HTMLElement | null;

  if (mainContent) {
    mainContent.style.marginTop = "";
  }

  ReactDom.unmountComponentAtNode(this.domElement);
}

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  private _escapeODataValue(value: string): string {
    return value.replace(/'/g, "''");
  }

  private _selectImageFromComputer(): void {
    const input = document.createElement("input");

    input.type = "file";
    input.accept = ".jpg,.jpeg,.png,.webp,image/jpeg,image/png,image/webp";

    input.addEventListener("change", () => {
      const file = input.files && input.files.length > 0
        ? input.files[0]
        : undefined;

      if (!file) {
        return;
      }

      this._uploadImage(file)
        .then((imageUrl: string) => {
          this.properties.backgroundImageUrl = imageUrl;
          this.context.propertyPane.refresh();
          this.render();
        })
        .catch((error: unknown) => {
          console.error("Site Hero image upload failed", error);

          window.alert(
            "העלאת התמונה נכשלה. פרטים נוספים מופיעים ב־Console."
          );
        });
    });

    input.click();
  }

  private async _uploadImage(file: File): Promise<string> {
    const webUrl = this.context.pageContext.web.absoluteUrl;

    const webRelativeUrl =
      this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, "");

    const folderRelativeUrl = `${webRelativeUrl}/SiteAssets`;

    const extensionIndex = file.name.lastIndexOf(".");
    const extension = extensionIndex >= 0
      ? file.name.substring(extensionIndex).toLowerCase()
      : "";

    const allowedExtensions = [".jpg", ".jpeg", ".png", ".webp"];

    if (allowedExtensions.indexOf(extension) === -1) {
      throw new Error("Unsupported image type.");
    }

    const fileName = `site-hero-${Date.now()}${extension}`;
    const safeFolderUrl = this._escapeODataValue(folderRelativeUrl);
    const safeFileName = this._escapeODataValue(fileName);

    const uploadUrl =
      `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${safeFolderUrl}')` +
      `/Files/add(url='${safeFileName}',overwrite=true)`;

    const uploadResponse: SPHttpClientResponse =
      await this.context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata.metadata=none"
          },
          body: file as unknown as string
        }
      );

    if (!uploadResponse.ok) {
      const errorText = await uploadResponse.text();

      throw new Error(
        `Image upload failed: ${uploadResponse.status} ${errorText}`
      );
    }

    return `${window.location.origin}${folderRelativeUrl}/${fileName}`;
  }

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    return {
      pages: [
        {
          header: {
            description: "הגדרות הבאנר"
          },
          groups: [
            {
              groupName: "תוכן הבאנר",
              groupFields: [
                PropertyPaneTextField("heroText", {
                  label: "כותרת ראשית"
                }),

                PropertyPaneTextField("subTitle", {
                  label: "כותרת משנה",
                  multiline: true
                }),

                PropertyPaneButton("uploadImageButton", {
                  text: this.properties.backgroundImageUrl
                    ? "החלפת תמונת הבאנר"
                    : "בחירת תמונה מהמחשב",
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._selectImageFromComputer.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}