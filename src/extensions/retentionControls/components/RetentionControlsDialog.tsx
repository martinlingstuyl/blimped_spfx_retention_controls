import { BaseDialog, IDialogConfiguration } from "@microsoft/sp-dialog";
import { BaseComponentContext } from "@microsoft/sp-component-base"
import * as ReactDOM from "react-dom";
import RetentionControlsDialogContent from "./RetentionControlsDialogContent";
import * as React from "react";

export default class RetentionControlsDialog extends BaseDialog {
    
    constructor(private context: BaseComponentContext, private listId: string, private listItemIds: number[]) {
        super();
    }
    
    public render(): void {        
        ReactDOM.render(<RetentionControlsDialogContent context={this.context} listId={this.listId} listItemIds={this.listItemIds} close={this.close}/>, this.domElement);
        //ReactDOM.render(<div>Test {this.context.pageContext.list?.id} {this.listId} {this.listItemIds.map(x => <div>{x}</div>)}</div>, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return { isBlocking: true };
    }

    protected onAfterClose(): void {
        super.onAfterClose();

        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}