import * as React from "react";
import styles from "./CrudReact.module.scss";
import { ICrudReactProps } from "./ICrudReactProps";
import { ICrudReactState } from "./ICrudReactState";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp, IListAddResult } from "@pnp/sp/presets/all";

export default class CrudReact extends React.Component<
  ICrudReactProps,
  ICrudReactState
> {
  constructor(props: ICrudReactProps, state: ICrudReactState) {
    super(props);

    sp.setup({
      spfxContext: this.props.context,
    });

    this.state = {
      Temp: "",

      MsgInsert: "",
      MsgEdit: "",
      MsgDelete: "",
      msgRead: "",
      msgReadAll: "",

      TitleInsert: "",
      EmailInsert: "",
      TitleEdit: "",
      EmailEdit: "",
      IDEdit: "",
      IDRead: "",
      IDDelete: "",
    };
  }
  public render(): React.ReactElement<ICrudReactProps> {
    return (
      <div className={styles.crudReact}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <p>
                <span className={styles.subTitle}>Insert</span>
              </p>
              <p>
                <span className={styles.label}>Nome</span>
                <input
                  type="text"
                  name="TitleInsert"
                  onChange={this.onChangeInput}
                />
              </p>
              <p>
                <span className={styles.label}>E-mail</span>
                <input
                  type="text"
                  name="EmailInsert"
                  onChange={this.onChangeInput}
                />
              </p>

              <p>
                <button
                  className={styles.button}
                  onClick={() => this.adcionarItem()}
                >
                  <span className={styles.label}>Criar item</span>
                </button>
              </p>
              <p>{this.state.MsgInsert}</p>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <p>
                <span className={styles.subTitle}>Ler 1</span>
              </p>
              <p>
                <span className={styles.label}>ID</span>
                <input
                  type="text"
                  name="IDRead"
                  onChange={this.onChangeInput}
                />
              </p>

              <p>
                <button
                  className={styles.button}
                  onClick={() => this.lerItem()}
                >
                  <span className={styles.label}>Ler item</span>
                </button>
              </p>
              <p>
                <ul
                  dangerouslySetInnerHTML={{ __html: this.state.msgRead }}
                ></ul>
              </p>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <p>
                <span className={styles.subTitle}>Ler Todos</span>
              </p>

              <p>
                <button
                  className={styles.button}
                  onClick={() => this.lerTodos()}
                >
                  <span className={styles.label}>Ler todos os items</span>
                </button>
              </p>
              <p>
                {" "}
                <ul
                  dangerouslySetInnerHTML={{ __html: this.state.msgReadAll }}
                ></ul>
              </p>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <p>
                <span className={styles.subTitle}>Editar Item</span>
              </p>
              <p>
                <span className={styles.label}>ID</span>
                <input
                  type="text"
                  name="IDEdit"
                  onChange={this.onChangeInput}
                />
              </p>
              <p>
                <span className={styles.label}>Nome</span>
                <input
                  type="text"
                  name="TitleEdit"
                  onChange={this.onChangeInput}
                />
              </p>
              <p>
                <span className={styles.label}>E-mail</span>
                <input
                  type="text"
                  name="EmailEdit"
                  onChange={this.onChangeInput}
                />
              </p>

              <p>
                <button
                  className={styles.button}
                  onClick={() => this.editItem()}
                >
                  <span className={styles.label}>Editar Item</span>
                </button>
              </p>
              <p>{this.state.MsgEdit}</p>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <p>
                <span className={styles.subTitle}>Delete </span>
              </p>
              <p>
                <span className={styles.label}>ID</span>
                <input
                  type="text"
                  name="IDDelete"
                  onChange={this.onChangeInput}
                />
              </p>

              <p>
                <button
                  className={styles.button}
                  onClick={() => this.deleteItem()}
                >
                  <span className={styles.label}>Criar item</span>
                </button>
              </p>
              <p>{this.state.MsgDelete}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
  onChangeInput = (event: React.ChangeEvent<HTMLInputElement>) => {
    const newState = { ...this.state };
    newState[event.target.name as keyof ICrudReactState] = event.target.value;
    this.setState(newState);
  };

  private adcionarItem() {
    sp.web.lists
      .getByTitle(this.props.listName)
      .items.add({
        Title: this.state.TitleInsert,
        Email: this.state.EmailInsert,
      })
      .then((i) => {
        console.log(
          this.setState({
            MsgInsert: `O Item ${i.data.ID} foi adicionado com sucesso.`,
          })
        );
      });
  }

  private lerItem() {
    sp.web.lists
      .getByTitle(this.props.listName)
      .items.getById(Number(this.state.IDRead))
      .get()
      .then((x) => {
        this.setState({
          msgRead: `<li>ID: ${x.ID}</li><li>Title: ${x.Title}</li><li>Email: ${x.Email}</li>`,
        });
      })
      .catch((x) => {
        console.log("teste");
        console.log(x);
        this.setState({
          msgRead: `Erro: O item ${this.state.IDRead} não foi encontrado. `,
        });
      });
  }

  private lerTodos() {
    sp.web.lists
      .getByTitle(this.props.listName)
      .items.getAll()
      .then((x) => {
        let tempstring = `Foram lidos ${x.length} itens`;
        x.forEach((element) => {
          tempstring += `<li>ID: ${element.ID} | Title: ${element.Title} | Email: ${element.Email}</li>`;
        });
        this.setState({
          msgReadAll: tempstring,
        });
      })
      .catch((x) => {
        console.log("teste");
        console.log(x);
        this.setState({
          msgReadAll: `Erro: Não foi possivel carregar os itens `,
        });
      });
  }

  private editItem() {
    sp.web.lists
      .getByTitle(this.props.listName)
      .items.getById(Number(this.state.IDEdit))
      .update({ Title: this.state.TitleEdit, Email: this.state.EmailEdit })
      .then((x) => {
        this.setState({
          MsgEdit: `O Item ${this.state.IDEdit} foi editado com sucesso.`,
        });
      })
      .catch((x) => {
        console.log("teste");
        console.log(x);
        this.setState({
          MsgEdit: `Erro: O item ${this.state.IDEdit} não foi encontrado. `,
        });
      });
  }

  private deleteItem() {
    if (!window.confirm("Você quer excluir o item ? ")) {
      return;
    }
    sp.web.lists
      .getByTitle(this.props.listName)
      .items.getById(Number(this.state.IDDelete))
      .delete()
      .then((x) => {
        this.setState({
          MsgDelete: `O item ${this.state.IDDelete} foi excluido com sucesso.`,
        });
      })
      .catch((x) => {
        console.log("teste");
        console.log(x);
        this.setState({
          MsgDelete: `Erro: O item ${this.state.IDRead} não foi encontrado. `,
        });
      });
  }
}
