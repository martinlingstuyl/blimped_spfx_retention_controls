import { Log } from "@microsoft/sp-core-library";
import { BaseListViewCommandSet, RowAccessor, type Command, type IListViewCommandSetExecuteEventParameters, type ListViewStateChangedEventArgs } from "@microsoft/sp-listview-extensibility";
import RetentionControlsDialog from "./components/RetentionControlsDialog";

export interface IRetentionControlsCommandSetProperties {}

const LOG_SOURCE: string = "RetentionControlsCommandSet";

export default class RetentionControlsCommandSet extends BaseListViewCommandSet<IRetentionControlsCommandSetProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized RetentionControlsCommandSet");

    const themePrimary = (window as unknown as { __themeState__: { theme: { themePrimary: string } } }).__themeState__.theme.themePrimary;
    const color = encodeURIComponent(this.context.isServedFromLocalhost ? "#ff0000" : themePrimary); //"#ff0000"

    const command: Command = this.tryGetCommand("RETENTION_CONTROLS_COMMAND");
    command.visible = false;
    command.iconImageUrl = `data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048" class="svg_dd790ee3" focusable="false"><path fill="${color}" d="M2048 128v640h-128v1152H128V768H0V128h2048zm-256 1664V768H256v1024h1536zm128-1152V256H128v384h1792zm-512 512H640v-128h768v128z"></path></svg>`;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  private openRetentionControls = async (listItems: readonly RowAccessor[]): Promise<void> => {
    if (!this.context.pageContext.list?.id) {
      return;
    }

    const dialog = new RetentionControlsDialog(this.context, this.context.pageContext.list.id.toString(), listItems);
    await dialog.show();
  };

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== "RETENTION_CONTROLS_COMMAND") {
      throw new Error("Unknown command");
    }

    this.openRetentionControls(event.selectedRows).catch(console.error);
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, "List view state changed");

    const command: Command = this.tryGetCommand("RETENTION_CONTROLS_COMMAND");
    if (command) {
      const hasSelectedItemsWithRetentionLabel = (this.context.listView.selectedRows && this.context.listView.selectedRows.length > 0 && this.context.listView.selectedRows?.some((row) => row.getValueByName("_ComplianceTag") !== undefined && row.getValueByName("_ComplianceTag") !== "")) || false;
      command.visible = hasSelectedItemsWithRetentionLabel;
    }

    this.raiseOnChange();
  };
}
