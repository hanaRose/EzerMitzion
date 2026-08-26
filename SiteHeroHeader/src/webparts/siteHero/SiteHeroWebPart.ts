 import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import SiteHero from "./components/SiteHero";
import { ISiteHeroProps } from "./components/ISiteHeroProps";

export interface ISiteHeroWebPartProps {
  listName: string;
  overrideHeroText: string;
  overrideSubTitle: string;
}

export default class SiteHeroWebPart extends BaseClientSideWebPart<ISiteHeroWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ISiteHeroProps> = React.createElement(
      SiteHero,
      {
        context: this.context,
        listName: this.properties.listName || "SiteHeader",
        overrideHeroText: this.properties.overrideHeroText,
        overrideSubTitle: this.properties.overrideSubTitle
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Site header settings"
          },
          groups: [
            {
              groupName: "מקור נתונים",
              groupFields: [
                PropertyPaneTextField("listName", {
                  label: "שם רשימה"
                })
              ]
            },
            {
              groupName: "עריכה ידנית (ידרוס את הרשימה)",
              groupFields: [
                PropertyPaneTextField("overrideHeroText", {
                  label: "כותרת ראשית",
                  placeholder: "השאר ריק לשימוש מהרשימה"
                }),
                PropertyPaneTextField("overrideSubTitle", {
                  label: "כותרת משנה",
                  placeholder: "השאר ריק לשימוש מהרשימה"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}