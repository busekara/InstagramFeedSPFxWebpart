import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";

import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "InstagramFeedWebPartStrings";
import InstagramFeed from "./components/InstagramFeed";
import { IInstagramFeedProps } from "./components/IInstagramFeedProps";

export interface IInstagramFeedWebPartProps {
  userToken: string;
  showIcon: Boolean;
  userFullName: string;
  accountName: string;
  layoutOneThirdRight: Boolean;
}

export default class InstagramFeedWebPart extends BaseClientSideWebPart<IInstagramFeedWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IInstagramFeedProps> =
      React.createElement(InstagramFeed, {
        userToken: this.properties.userToken,
        showIcon: this.properties.showIcon,
        userFullName: this.properties.userFullName,
        accountName: this.properties.accountName,
        layoutOneThirdRight: this.properties.layoutOneThirdRight,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField("userToken", {
                  label: strings.UsertokenFieldLabel
                    ? strings.UsertokenFieldLabel
                    : strings.DefaultUsertoken,
                }),
                PropertyPaneTextField("userFullName", {
                  label: "User Full Name :",
                }),
                PropertyPaneTextField("accountName", {
                  label: "Account Name : ",
                }),
                PropertyPaneToggle("showIcon", {
                  label: strings.ShowIconToggleLabel,
                  onText: strings.ShowIconToggleTrueLabel,
                  offText: strings.ShowIconToggleFalseLabel,
                }),
                PropertyPaneToggle("layoutOneThirdRight", {
                  label: "Layout One-Third Right",
                  onText: "open",
                  offText: "close",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
