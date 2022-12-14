import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-webpart-base";

import * as strings from "MyTodoListWebPartWebPartStrings";
import MyTodoListWebPart from "./components/MyTodoListWebPart";
import { IMyTodoListWebPartProps } from "./components/IMyTodoListWebPartProps";
import { sp } from "@pnp/sp";

export interface IMyTodoListWebPartWebPartProps {
  description: string;
  context: any;
}
export default class MyTodoListWebPartWebPart extends BaseClientSideWebPart<IMyTodoListWebPartWebPartProps> {
  public onInit() {
    return new Promise<void>((resolve, _reject) => {
      sp.setup({
        spfxContext: this.context,
      });
      resolve(undefined);
    });
  }
  public render(): void {
    const element: React.ReactElement<IMyTodoListWebPartProps> =
      React.createElement(MyTodoListWebPart, {
        description: this.properties.description,
        context: this.context,
      });

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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
