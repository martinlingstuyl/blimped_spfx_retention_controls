import { BaseComponentContext } from "@microsoft/sp-component-base";
import * as React from "react";
import { useState, useEffect } from "react";
import { SharePointService } from "../../../shared/services/SharePointService";
import { PermissionsService } from "../../../shared/services/PermissionsService";
import { RowAccessor } from "@microsoft/sp-listview-extensibility";
import { IDriveItem } from "../../../shared/interfaces/IDriveItem";
import { SingleItemView } from "./SingleItemView";
import { MultiItemView } from "./MultiItemView";
import { IBatchItemResponse } from "../../../shared/interfaces/IBatchErrorResponse";
import { LibraryView } from "./LibraryView";
import { IPagedDriveItems } from "../../../shared/interfaces/IPagedDriveItems";
import { IPermissions } from "../../../shared/interfaces/IPermissions";

export interface IRetentionControlsDialogProps {
  context: BaseComponentContext;
  listId: string;
  listItems: readonly RowAccessor[];
  selectedItems: number;
  permissions?: IPermissions;
  onClose: { (): void };
}

const RetentionControlsDialog: React.FC<IRetentionControlsDialogProps> = (props) => {
  const { selectedItems } = props;
  const spoService = props.context.serviceScope.consume(SharePointService.serviceKey);
  const permissionsService = props.context.serviceScope.consume(PermissionsService.serviceKey);

  const [hasEditPermissions, setHasEditPermissions] = useState<boolean>(false);

  useEffect(() => {
    const checkPermissions = async (): Promise<void> => {
      if (props.permissions && props.permissions.entries && props.permissions.entries.length > 0) {
        const userPermissions = await permissionsService.getUserPermissions(props.permissions.entries);
        setHasEditPermissions(userPermissions.editMode && userPermissions.hasSharePointEditPermissions);
      } else {
        // If no permissions are configured, check SharePoint permissions only
        const userPermissions = await permissionsService.getUserPermissions([]);
        setHasEditPermissions(userPermissions.hasSharePointEditPermissions);
      }
    };

    checkPermissions().catch(console.error);
  }, [props.permissions]);

  const fetchItemMetadata = async (listItemIds: number[]): Promise<IDriveItem[]> => {            
    return await spoService.getDriveItems(props.listId, listItemIds);
  };

  const fetchItemsPaged = async (pageSize: number, nextLink?: string): Promise<IPagedDriveItems> => {                
    return nextLink !== undefined ?
      await spoService.getPagedDriveItemsUsingNextLink(nextLink) :
      await spoService.getPagedDriveItems(props.listId, pageSize);
  };

  const onClearingLabels = async (listItemIds: number[]): Promise<IBatchItemResponse[]> => {
    if (!hasEditPermissions || listItemIds.length === 0) {
      return [];
    }

    return await spoService.clearRetentionLabels(listItemIds);        
  };

  const onTogglingRecords = async (listItemIds: number[], newLockState: boolean): Promise<IBatchItemResponse[]> => {
    if (!hasEditPermissions || listItemIds.length === 0) {
      return [];
    }

    return await spoService.toggleLockStatus(listItemIds, newLockState);        
  };

  return selectedItems === 1 ? <>
    <SingleItemView isServedFromLocalhost={props.context.isServedFromLocalhost} listItems={props.listItems} hasEditPermissions={hasEditPermissions} onFetching={fetchItemMetadata} onClearing={onClearingLabels} onToggling={onTogglingRecords} onClose={props.onClose} />          
  </> : selectedItems > 1 ? <>
    <MultiItemView isServedFromLocalhost={props.context.isServedFromLocalhost} listItems={props.listItems} hasEditPermissions={hasEditPermissions} onFetching={fetchItemMetadata} onClearing={onClearingLabels} onToggling={onTogglingRecords} onClose={props.onClose} />
  </> : <>
    <LibraryView isServedFromLocalhost={props.context.isServedFromLocalhost} hasEditPermissions={hasEditPermissions} onFetching={fetchItemMetadata} onFetchingPaged={fetchItemsPaged} onClearing={onClearingLabels} onToggling={onTogglingRecords} onClose={props.onClose} />
  </>;
};

export default RetentionControlsDialog;
