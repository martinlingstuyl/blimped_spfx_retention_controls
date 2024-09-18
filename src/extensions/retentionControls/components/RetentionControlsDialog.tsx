import { BaseComponentContext } from "@microsoft/sp-component-base";
import * as React from "react";
import { SharePointService } from "../../../shared/services/SharePointService";
import { RowAccessor } from "@microsoft/sp-listview-extensibility";
import { IDriveItem } from "../../../shared/interfaces/IDriveItem";
import { SingleItemView } from "./SingleItemView";
import { MultiItemView } from "./MultiItemView";
import { IBatchItemResponse } from "../../../shared/interfaces/IBatchErrorResponse";
import { LibraryView } from "./LibraryView";
import { IPagedDriveItems } from "../../../shared/interfaces/IPagedDriveItems";

export interface IRetentionControlsDialogProps {
  context: BaseComponentContext;
  listId: string;
  listItems: readonly RowAccessor[];
  selectedItems: number;
  onClose: { (): void };
}

const RetentionControlsDialog: React.FC<IRetentionControlsDialogProps> = (props) => {
  const { selectedItems } = props;
  const spoService = props.context.serviceScope.consume(SharePointService.serviceKey);  

  const fetchItemMetadata = async (listItemIds: number[]): Promise<IDriveItem[]> => {            
    return await spoService.getDriveItems(props.listId, listItemIds);
  };

  const fetchItemsPaged = async (pageSize: number, nextLink?: string): Promise<IPagedDriveItems> => {                
    return nextLink !== undefined ?
      await spoService.getPagedDriveItemsUsingNextLink(nextLink) :
      await spoService.getPagedDriveItems(props.listId, pageSize);
  };

  const onClearingLabels = async (listItemIds: number[]): Promise<IBatchItemResponse[]> => {
    if (listItemIds.length === 0) {
      return [];
    }

    return await spoService.clearRetentionLabels(listItemIds);        
  };

  const onTogglingRecords = async (listItemIds: number[], newLockState: boolean): Promise<IBatchItemResponse[]> => {
    if (listItemIds.length === 0) {
      return [];
    }

    return await spoService.toggleLockStatus(listItemIds, newLockState);        
  };

  return selectedItems === 1 ? <>
    <SingleItemView listItems={props.listItems} onFetching={fetchItemMetadata} onClearing={onClearingLabels} onToggling={onTogglingRecords} onClose={props.onClose} />          
  </> : selectedItems > 1 ? <>
    <MultiItemView listItems={props.listItems} onFetching={fetchItemMetadata} onClearing={onClearingLabels} onToggling={onTogglingRecords} onClose={props.onClose} />
  </> : <>
    <LibraryView onFetching={fetchItemMetadata} onFetchingPaged={fetchItemsPaged} onClearing={onClearingLabels} onToggling={onTogglingRecords} onClose={props.onClose} />
  </>;
};

export default RetentionControlsDialog;
