import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import * as strings from "ArticleDisplayWebPartStrings";
import ArticleDisplay from "./components/ArticleDisplay";
import { IArticleDisplayProps } from "./components/IArticleDisplayProps";

export interface IArticleDisplayWebPartProps {
  description: string;
  multiSelect: string[];
  selectedGroupId: string;
  selectedView: string;
}

export interface IView {
  Id: string;
  Title: string;
}

export default class ArticleDisplayWebPart extends BaseClientSideWebPart<IArticleDisplayWebPartProps> {
  private termSetOptions: IPropertyPaneDropdownOption[] = [];
  private termStoreGroupOptions: IPropertyPaneDropdownOption[] = [];
  private viewOptions: IView[] = [];
  public render(): void {
    const element: React.ReactElement<IArticleDisplayProps> =
      React.createElement(ArticleDisplay, {
        context: this.context,
        groupId: this.properties.selectedGroupId,
        setNames: this.properties.multiSelect,
        selectedView: this.properties.selectedView,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this.getTermStoreGroups();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public async getViews(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Site Pages')/views`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch views. Error code: ${response.status}`);
    }

    const data = await response.json();
    this.viewOptions = [
      {
        Id: "",
        Title: "Select Set",
      },
      ...data.value.map((view: any) => ({
        Id: view.Id,
        Title: view.Title,
      })),
    ];
  }
  private async getTermSets(groupId: string): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/v2.1/termstore/groups/${groupId}/sets`;
    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();

    // Populate the dropdown options
    this.termSetOptions = [
      {
        key: "",
        text: "Select Set",
      },
      ...data.value.map((termSet: any) => {
        return {
          key: termSet.localizedNames[0].name,
          text: termSet.localizedNames[0].name,
        };
      }),
    ];
  }

  private async getTermStoreGroups(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/v2.1/termstore/groups`;
    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    await this.getViews();
    this.termStoreGroupOptions = data.value.map((group: any) => {
      return {
        key: group.id,
        text: group.name,
      };
    });
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    // Load the term store groups if not already loaded
    if (this.termStoreGroupOptions.length === 0) {
      await this.getTermStoreGroups();
    }

    if (this.termSetOptions.length === 0 && this.properties.selectedGroupId) {
      await this.getTermSets(this.properties.selectedGroupId);
    }

    this.context.propertyPane.refresh();
  }

  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): Promise<void> {
    if (propertyPath === "selectedGroupId" && newValue) {
      await this.getTermSets(newValue);
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown("selectedGroupId", {
                  label: "Select Term Store Group",
                  options: this.termStoreGroupOptions,
                  selectedKey: this.properties.selectedGroupId,
                }),
                PropertyFieldMultiSelect("multiSelect", {
                  key: "multiSelect",
                  label: "Multi select field",
                  options: this.termSetOptions,
                  selectedKeys: this.properties.multiSelect,
                }),
                PropertyPaneDropdown("selectedView", {
                  label: "Select View",
                  options: this.viewOptions.map((view) => ({
                    key: view.Id,
                    text: view.Title,
                  })),
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
