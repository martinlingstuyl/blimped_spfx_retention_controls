import * as ReactDOM from "react-dom";
import * as React from "react";
import Dialog, { DialogFooter, DialogType } from "@fluentui/react/lib/Dialog";
import { DefaultButton } from "@fluentui/react/lib/Button";
import * as strings from "RetentionControlsCommandSetStrings";

export default class AlertDialogManager {
  private domElement: HTMLDivElement | null = null;
  private onClosedCallback: (confirmed?: boolean) => void;

  constructor(private title: string, private subText: string, private buttonText?: string) {
  }

  public async close(confirmed?: boolean): Promise<void> {
    if (this.domElement) {
      ReactDOM.unmountComponentAtNode(this.domElement);
      this.domElement.remove();
      this.domElement = null;
    }

    if (this.onClosedCallback !== undefined) {
      this.onClosedCallback(confirmed);
    }
  }

  public onClosed(callback: (confirmed?:boolean)=>void): void {
    this.onClosedCallback = callback;
  }

  public async show(): Promise<void> {
    this.domElement = document.createElement('div');
    document.body.appendChild(this.domElement);

    const close = async (confirmed?: boolean): Promise<void> => {
        await this.close(confirmed);
    };

    const reactElement = <Dialog
      hidden={false}
      onDismiss={() => close()}      
      dialogContentProps={{
        type: DialogType.normal,
        showCloseButton: false,
        title: this.title,
        subText: this.subText,
      }}>
      <DialogFooter>        
        { <DefaultButton onClick={() => close(false)} text={this.buttonText ?? strings.CloseModal} /> }
      </DialogFooter>
    </Dialog>

    ReactDOM.render(reactElement, 
      this.domElement);
  }
}
