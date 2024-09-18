import { BaseComponentContext } from "@microsoft/sp-component-base";
import * as ReactDOM from "react-dom";
import RetentionControlsDialog from "./components/RetentionControlsDialog";
import * as React from "react";
import { RowAccessor } from "@microsoft/sp-listview-extensibility";

export default class RetentionControlsDialogManager {
  private domElement: HTMLDivElement | null = null;

  constructor(private context: BaseComponentContext, private listId: string, private listItems: readonly RowAccessor[], private selectedItems: number) {
  }

  public async close(): Promise<void> {
    if (this.domElement) {
      ReactDOM.unmountComponentAtNode(this.domElement);
      this.domElement.remove();
      this.domElement = null;
    }
  }

  public async show(): Promise<void> {
    this.domElement = document.createElement('div');
    document.body.appendChild(this.domElement);

    const close = async (): Promise<void> => {
        await this.close();
    };

    ReactDOM.render(<RetentionControlsDialog context={this.context} listId={this.listId} listItems={this.listItems} selectedItems={this.selectedItems} onClose={close} />, 
      this.domElement);
  }
}
