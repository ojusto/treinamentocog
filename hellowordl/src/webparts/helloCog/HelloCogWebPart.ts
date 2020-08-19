import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./HelloCogWebPart.module.scss";
import * as strings from "HelloCogWebPartStrings";

export interface IHelloCogWebPartProps {
  description: string;
}

export default class HelloCogWebPart extends BaseClientSideWebPart<
  IHelloCogWebPartProps
> {
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloCog}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Ola Mundo!</span>
              <p class="${styles.subTitle}">Essa é nossa primeira webpart.</p>
              <p class="${styles.description}">${escape(
      this.properties.description
    )}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
