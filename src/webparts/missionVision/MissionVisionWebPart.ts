import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { PropertyPaneDropdown } from "@microsoft/sp-property-pane";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import MissionVision from "./components/MissionVision";
import { IMissionVisionProps } from "./components/IMissionVisionProps";

export interface IMissionVisionWebPartProps {
  listName: string;
  gridBackground: string;
  gridGap: string;
  itemAsLink: boolean;
  itemWidth: string;
  itemBackground: string;
  iconWidth: string;
  titleFontSize: string;
  descFontSize: string;
  alignment: string;
}

export default class MissionVisionWebPart extends BaseClientSideWebPart<IMissionVisionWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IMissionVisionProps> =
      React.createElement(MissionVision, {
        listName: this.properties.listName,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        alignment: this.properties.alignment,
        gridBackground: this.properties.gridBackground,
        gridGap: this.properties.gridGap,
        itemAsLink: this.properties.itemAsLink,
        itemWidth: this.properties.itemWidth,
        itemBackground: this.properties.itemBackground,
        iconWidth: this.properties.iconWidth,
        titleFontSize: this.properties.titleFontSize,
        descFontSize: this.properties.descFontSize,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onInit(): Promise<void> {
    if (this.properties.itemAsLink === undefined) {
      this.properties.itemAsLink = true;
    }
    if (this.properties.alignment === undefined) {
      this.properties.alignment = "center";
    }
    return super.onInit();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Mission Vision Web Part Settings",
          },
          groups: [
            {
              groupName: "List Settings",
              groupFields: [
                PropertyPaneTextField("listName", {
                  label: "SharePoint List Name",
                  description:
                    "Enter the exact name of the SharePoint list to pull grid items from. " +
                    "The list must contain these columns: Title, Description, URL. " +
                    "The icon image should be attached to each item as an attachment.",
                  value: this.properties.listName,
                }),
              ],
            },
            {
              groupName: "Grid Style",
              groupFields: [
                PropertyPaneDropdown("alignment", {
                  label: "Grid Alignment",
                  options: [
                    { key: "left", text: "Left" },
                    { key: "center", text: "Center" },
                    { key: "right", text: "Right" },
                  ],
                  selectedKey: this.properties.alignment,
                }),
                PropertyPaneTextField("gridBackground", {
                  label: "Grid Background",
                  value: this.properties.gridBackground,
                }),
                PropertyPaneTextField("gridGap", {
                  label: "Grid Gap",
                  value: this.properties.gridGap,
                }),
                PropertyPaneCheckbox("itemAsLink", {
                  text: "Use item as a link",
                  checked: this.properties.itemAsLink,
                }),
              ],
            },
            {
              groupName: "Item Style",
              groupFields: [
                PropertyPaneTextField("itemWidth", {
                  label: "Item Width",
                  value: this.properties.itemWidth,
                }),
                PropertyPaneTextField("itemBackground", {
                  label: "Item Background",
                  value: this.properties.itemBackground,
                }),
                PropertyPaneTextField("iconWidth", {
                  label: "Icon Size",
                  value: this.properties.iconWidth,
                }),
                PropertyPaneTextField("titleFontSize", {
                  label: "Title Font Size",
                  value: this.properties.titleFontSize,
                }),
                PropertyPaneTextField("descFontSize", {
                  label: "Description Font Size",
                  value: this.properties.descFontSize,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
