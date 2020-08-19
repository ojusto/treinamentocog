import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./CrudSemFreameworkWebPart.module.scss";
import * as strings from "CrudSemFreameworkWebPartStrings";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface ICrudSemFreameworkWebPartProps {
  listName: string;
}

interface IListItem {
  Title?: string;
  Email?: string;
  Id: number;
}

export default class CrudSemFreameworkWebPart extends BaseClientSideWebPart<
  ICrudSemFreameworkWebPartProps
> {
  private listItemEntityTypeName: string = undefined;

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.crudSemFreamework}">
    <div class="${styles.container}">
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <span class="ms-font-xl ms-fontColor-white  ${styles.title}">
            Crud sem FrameWork
          </span>
        </div>
      </div>
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">

      <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
        <p><span class="ms-font-xl ms-fontColor-white ${styles.subTitle}">
          Insert
        </span></p>
        <p><span class="${styles.label}">Nome</span>
        <input type="text" class="txtTitulo" id="txtTitulo" ></p>
        <p><span class="${styles.label}">E-mail</span>
        <input type="text" class="txtEmail" id="txtEmail" ></p>

        <p><button class="${styles.button} create-Button">
          <span class="${styles.label}">Criar Item</span>
        </button><p>

      </div>
      <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <div class="statusInsert"></div>
          <ul class="itemsInsert"><ul>
        </div>
    </div>
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <p><span class="ms-font-xl ms-fontColor-white ${styles.subTitle}">
          Ler 1
          </span></p>
          <p><span class="${styles.label}">ID</span>
          <input type="text" class="txtIDedit" id="txtIDLer" ></p>
          <button class="${styles.button} read-Button">
            <span class="${styles.label}">Ler 1 Item</span>
          </button>
        </div>
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <div class="statusReadOne"></div>
          <ul class="itemsReadOne"><ul>
        </div>
      </div>
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
        <p><span class="ms-font-xl ms-fontColor-white ${styles.subTitle}">
          Let todos
          </span></p>
          <button class="${styles.button} readall-Button">
            <span class="${styles.label}">Ler todos os items</span>
          </button>
        </div>
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <div class="statusReadAll"></div>
          <ul class="itemsReadAll"><ul>
        </div>
      </div>
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
        <p><span class="ms-font-xl ms-fontColor-white ${styles.subTitle}">
          Edit
          </span></p>
          <p><span class="${styles.label}">ID</span>
          <input type="text" class="txtIDedit" id="txtIDedit" ></p>
          <p><span class="${styles.label}">Nome</span>
          <input type="text" class="txtTitulo" id="txtTituloedit" ></p>
          <p><span class="${styles.label}">E-mail</span>
          <input type="text" class="txtEmail" id="txtEmailedit" ></p>
          <button class="${styles.button} update-Button">
            <span class="${styles.label}">Update item</span>
          </button>

        </div>
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <div class="statusEdit"></div>
          <ul class="itemsEdit"><ul>
        </div>
      </div>
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
        <p><span class="ms-font-xl ms-fontColor-white ${styles.subTitle}">
          Delete
          </span></p>
          <p><span class="${styles.label}">ID</span>
          <input type="text" class="txtIDedit" id="txtIDDelete" ></p>
          <button class="${styles.button} delete-Button">
            <span class="${styles.label}">Delete item</span>
          </button>
        </div>
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <div class="statusDelete"></div>
          <ul class="itemsDelete"><ul>
        </div>
      </div>

    </div>
  </div>`;

    this.setButtonsEventHandlers();
  }
  private setButtonsEventHandlers(): void {
    const webPart: CrudSemFreameworkWebPart = this;
    this.domElement
      .querySelector("button.create-Button")
      .addEventListener("click", () => {
        webPart.createItem();
      });
    this.domElement
      .querySelector("button.read-Button")
      .addEventListener("click", () => {
        webPart.readItem();
      });
    this.domElement
      .querySelector("button.readall-Button")
      .addEventListener("click", () => {
        webPart.readAllItems();
      });
    this.domElement
      .querySelector("button.update-Button")
      .addEventListener("click", () => {
        webPart.updateItem();
      });
    this.domElement
      .querySelector("button.delete-Button")
      .addEventListener("click", () => {
        webPart.deleteItem();
      });
  }

  private createItem(): void {
    this.getListItemEntityTypeName()
      .then(
        (listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
          const body: string = JSON.stringify({
            __metadata: {
              type: listItemEntityTypeName,
            },
            Title: (<HTMLInputElement>document.getElementById("txtTitulo"))
              .value,
            Email: (<HTMLInputElement>document.getElementById("txtEmail"))
              .value,
          });
          return this.context.spHttpClient.post(
            `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
              },
              body: body,
            }
          );
        }
      )
      .then(
        (response: SPHttpClientResponse): Promise<IListItem> => {
          return response.json();
        }
      )
      .then((item: IListItem): void => {
        this.updateStatus(
          `O item '${item.Title}' com o ID: ${item.Id}) foi adicionado.`,
          [],
          "Insert"
        );
      });
  }

  private readItem(): void {
    this.context.spHttpClient
      .get(
        `${
          this.context.pageContext.web.absoluteUrl
        }/_api/web/lists/getbytitle('${this.properties.listName}')/items(${
          (<HTMLInputElement>document.getElementById("txtIDLer")).value
        })?$select=Title,Email,Id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        }
      )
      .then(
        (response: SPHttpClientResponse): Promise<IListItem> => {
          return response.json();
        }
      )
      .then((item: IListItem): void => {
        if (item.Id === undefined) {
          this.updateStatus(
            `Não foi  encontrado nenhum item com este ID.`,
            [],
            "ReadOne"
          );
          throw new Error("Não foi  encontrado nenhum item com este ID.");
        }
        this.updateStatus(
          `Item ID: ${item.Id}, Title: ${item.Title} , E-mail: ${item.Email}`,
          [],
          "ReadOne"
        );
      });
  }

  private readAllItems(): void {
    this.updateStatus("Loading all items...", [], "ReadAll");
    this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$select=Title,Email,Id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        }
      )
      .then(
        (response: SPHttpClientResponse): Promise<{ value: IListItem[] }> => {
          return response.json();
        }
      )
      .then((response: { value: IListItem[] }): void => {
        this.updateStatus(
          `Foram carregados ${response.value.length} items`,
          response.value,
          "ReadAll"
        );
      });
  }

  private updateItem(): void {
    this.updateStatus("Loading latest items...", [], "Edit");
    let latestItemId: number = Number(
      (<HTMLInputElement>document.getElementById("txtIDedit")).value
    );
    this.getListItemEntityTypeName()
      .then(
        (listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
          const body: string = JSON.stringify({
            __metadata: {
              type: listItemEntityTypeName,
            },
            Title: (<HTMLInputElement>document.getElementById("txtTituloedit"))
              .value,
            Email: (<HTMLInputElement>document.getElementById("txtEmailedit"))
              .value,
          });
          return this.context.spHttpClient.post(
            `${
              this.context.pageContext.web.absoluteUrl
            }/_api/web/lists/getbytitle('${this.properties.listName}')/items(${
              (<HTMLInputElement>document.getElementById("txtIDedit")).value
            })`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",
              },
              body: body,
            }
          );
        }
      )
      .then((response: SPHttpClientResponse): void => {
        console.log(response);
        if (!response.ok) {
          this.updateStatus(
            `Não foi  encontrado nenhum item com este ID.`,
            [],
            "Edit"
          );
          throw new Error("Não foi  encontrado nenhum item com este ID.");
        } else {
          this.updateStatus(
            `O item ${latestItemId} foi alterado com sucesso`,
            [],
            "Edit"
          );
        }
      });
  }

  private deleteItem(): void {
    this.updateStatus("Loading latest items...", [], "Delete");
    let latestItemId: number = Number(
      (<HTMLInputElement>document.getElementById("txtIDedit")).value
    );
    if (!window.confirm(`Você deseja deletar o item ${latestItemId}?"`)) {
      return;
    }

    this.getListItemEntityTypeName()

      .then(
        (listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
          this.updateStatus(
            `Deleting item with ID: ${latestItemId}...`,
            [],
            "Delete"
          );
          return this.context.spHttpClient.post(
            `${
              this.context.pageContext.web.absoluteUrl
            }/_api/web/lists/getbytitle('${this.properties.listName}')/items(${
              (<HTMLInputElement>document.getElementById("txtIDDelete")).value
            })`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
                "IF-MATCH": "*",
                "X-HTTP-Method": "DELETE",
              },
            }
          );
        }
      )
      .then((response: SPHttpClientResponse): void => {
        if (!response.ok) {
          this.updateStatus(
            `Não foi  encontrado nenhum item com este ID.`,
            [],
            "Delete"
          );
          throw new Error("Não foi  encontrado nenhum item com este ID.");
        } else {
          this.updateStatus(
            `O item ${latestItemId} foi deletado.`,
            [],
            "Delete"
          );
        }
      });
  }

  private updateStatus(
    status: string,
    items: IListItem[] = [],
    loc: string
  ): void {
    this.domElement.querySelector(".status" + loc).innerHTML = status;
    this.updateItemsHtml(items, loc);
  }

  private getListItemEntityTypeName(): Promise<string> {
    return new Promise<string>(
      (
        resolve: (listItemEntityTypeName: string) => void,
        reject: (error: any) => void
      ): void => {
        if (this.listItemEntityTypeName) {
          resolve(this.listItemEntityTypeName);
          return;
        }

        this.context.spHttpClient
          .get(
            `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')?$select=ListItemEntityTypeFullName`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "odata-version": "",
              },
            }
          )
          .then(
            (
              response: SPHttpClientResponse
            ): Promise<{ ListItemEntityTypeFullName: string }> => {
              return response.json();
            },
            (error: any): void => {
              reject(error);
            }
          )
          .then((response: { ListItemEntityTypeFullName: string }): void => {
            this.listItemEntityTypeName = response.ListItemEntityTypeFullName;
            resolve(this.listItemEntityTypeName);
          });
      }
    );
  }

  private updateItemsHtml(items: IListItem[], loc: string): void {
    this.domElement.querySelector(".items" + loc).innerHTML = items
      .map((item) => `<li>${item.Title} - ${item.Email} - (${item.Id})</li>`)
      .join("");
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
                PropertyPaneTextField("listName", {
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
