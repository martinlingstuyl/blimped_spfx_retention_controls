import * as React from "react";
import { useEffect, useState } from "react";
import { IDriveItem } from "../../../shared/interfaces/IDriveItem";
import * as strings from "RetentionControlsCommandSetStrings";
import { RowAccessor } from "@microsoft/sp-listview-extensibility";
import { format, SelectionMode } from "@fluentui/react/lib/Utilities";
import { ShimmeredDetailsList } from "@fluentui/react/lib/ShimmeredDetailsList";
import { IItemMetadata } from "../../../shared/interfaces/IItemMetadata";
import { ICustomColumn } from "../../../shared/interfaces/ICustomColumn";
import { ItemColumn } from "./ItemColumn";
import Dialog, { DialogFooter, DialogType } from "@fluentui/react/lib/Dialog";
import { DefaultButton, IButtonProps, PrimaryButton } from "@fluentui/react/lib/Button";
import { dialogFooterStyles, messageBarStyles } from "../../../shared/styles";
import { Stack } from "@fluentui/react/lib/Stack";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { ContextualMenu, IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import { Log } from "@microsoft/sp-core-library";
import { LOG_SOURCE } from "../RetentionControlsCommandSet";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { IItemState } from "../../../shared/interfaces/IItemState";
import { IBatchItemResponse } from "../../../shared/interfaces/IBatchErrorResponse";
import { flattenItemMetadata, flattenItemMetadataList, updateObjectProperties } from "../../../shared/utils";
import { INotification } from "../../../shared/interfaces/INotification";
import { ResponsiveMode } from "@fluentui/react/lib/ResponsiveMode";
import { itemMetadataColumns } from "../../../shared/constants";

export interface IMultiItemView {
  listItems: readonly RowAccessor[];
  onClose: () => void;
  onFetching: (listItemIds: number[]) => Promise<IDriveItem[]>;
  onClearing: (listItemIds: number[]) => Promise<IBatchItemResponse[]>;
  onToggling: (listItemIds: number[], newLockstate: boolean) => Promise<IBatchItemResponse[]>;
}

export const MultiItemView: React.FC<IMultiItemView> = (props) => {
  const pageSize = 10;
  const shimmerLines = props.listItems.length > pageSize ? pageSize : props.listItems.length;  
  const [listItemIds, setListItemIds] = useState<number[]>(props.listItems.map(i => parseFloat(i.getValueByName("ID"))));
  const [loading, setLoading] = useState<boolean>(true);
  const [notification, setNotification] = useState<INotification | undefined>();
  const [pageNumber, setPageNumber] = useState<number>(1);  
  const [totalPages, setTotalPages] = useState<number>(1);  
  const [executingAction, setExecutingAction] = useState<boolean>(false);  
  const [actionStatus, setActionStatus] = useState<string>("");
  const [fetchedItems, setFetchedItems] = useState<IItemMetadata[]>([]);
  const [itemsState, setItemsState] = useState<IItemState[]>([]);
  const [itemsWithMetadata, setItemsWithMetadata] = useState<IItemMetadata[] | undefined>();
  
  const refreshItemMetadata = async (listItemId: number): Promise<void> => {
    if (itemsWithMetadata === undefined) {
      return;
    } 

    try {
        const response = await props.onFetching([listItemId]);    

        if (!response || response.length === 0) {
          throw new Error(strings.ErrorFetchingData);
        }

        const updatedItem = flattenItemMetadata(response[0]) as IItemMetadata;
        const item = itemsWithMetadata.filter(i => i.id === updatedItem.id)[0];
        
        updateObjectProperties(item, updatedItem);

        setItemsWithMetadata([...itemsWithMetadata]);
    } catch (error) {
      const message = strings.ErrorFetchingData + ": " + error.message;
      Log.error(LOG_SOURCE, new Error(message));
      setNotification({ message: message, notificationType: MessageBarType.error });
    }
  };

  const fetchAllItems = async (): Promise<IItemMetadata[]> => {
    const allItemsWithMetadata: IItemMetadata[] = [];
    const pages = Math.ceil(listItemIds.length / pageSize);

    for (let page = 1; page <= pages; page++) {
      const itemIds = listItemIds.slice((page - 1) * pageSize, page * pageSize);
      const response = await props.onFetching(itemIds);

      for (const item of flattenItemMetadataList(response)) {
        if (item.retentionLabel !== "") {
          allItemsWithMetadata.push(item);
        }
      }
    } 
    
    return allItemsWithMetadata;
  }

  const fetchData = async (data: number[], page: number, showLoader: boolean = true, appendToFetchedList: boolean = true): Promise<void> => {
    try {
      if (showLoader)
        setLoading(true);

      let fetchedItemsList = fetchedItems;
      
      if (!appendToFetchedList)
        fetchedItemsList = [];

      setPageNumber(page);

      if (data.length > 0) {
        const retrieveNewItems = (page === 1 && fetchedItemsList.length === 0) || (page * pageSize > (fetchedItemsList?.length ?? 0) && fetchedItemsList?.length > ((page-1) * pageSize));

        if (retrieveNewItems) {
          const itemIds = data.slice((page - 1) * pageSize, (page * pageSize) + 1); // Retrieve one item extra to check if more pages are available.
          const response = await props.onFetching(itemIds);
          const newItems = flattenItemMetadataList(response);
          
          for (const item of newItems) {
            if (!fetchedItemsList.some(i => i.driveItemId === item.driveItemId)) {
              fetchedItemsList.push(item);
            }
          }
          
          setFetchedItems([...fetchedItemsList]);
        }


        const items = fetchedItemsList.slice((page - 1) * pageSize, page * pageSize);
        setItemsWithMetadata(items);
        setTotalPages(Math.ceil(fetchedItemsList.length / pageSize));
      }
      else {
        setItemsWithMetadata([]);
        setTotalPages(1);
      }
      
      if (showLoader)
        setLoading(false);
    } catch (error) {
      const message = strings.ErrorFetchingData + ": " + error.message;
      Log.error(LOG_SOURCE, new Error(message));
      setNotification({ message: message, notificationType: MessageBarType.error });

      if (showLoader)
        setLoading(false);
    }
  };

  const onTogglingRecord = async (item: IItemMetadata): Promise<void> => {
    if (itemsState === undefined || itemsWithMetadata === undefined) {
      return;
    }

    Log.info(LOG_SOURCE, `Toggling record status for '${item.name}'`);

    setItemsState([...itemsState.filter(i => i.listItemId !== item.id), { listItemId: item.id, toggling: true, errorToggling: undefined, clearing: false, errorClearing: false }]);
    setNotification(undefined);    
    setExecutingAction(true);
    
    // Trigger re-render table
    setItemsWithMetadata([...itemsWithMetadata]);

    try {
      const newLockState = !item.isRecordLocked;
      const response = await props.onToggling([item.id], newLockState);
      const success = response[0].success;

      if (!success) {        
        setNotification({ message: format(strings.ToggleErrorForSingleItem, newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase()) + " " + response[0].errorMessage, notificationType: MessageBarType.error });
      }
      else {        
        setNotification({ message: format(strings.RecordStatusToggled, newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase()), notificationType: MessageBarType.success });
      }
      
      setItemsState([...itemsState.filter(i => i.listItemId !== item.id), { listItemId: item.id, toggling: false, errorToggling: response[0].errorMessage, clearing: false, errorClearing: false }]);
      await refreshItemMetadata(item.id);
    }
    catch (error) {
      setNotification({ message: error.message, notificationType: MessageBarType.error });
      Log.error(LOG_SOURCE, new Error(error.message));      

      setItemsState([...itemsState.filter(i => i.listItemId !== item.id), { listItemId: item.id, toggling: false, errorToggling: undefined, clearing: false, errorClearing: false }]);
      
      // Trigger re-render table
      setItemsWithMetadata([...itemsWithMetadata]);
    }
    finally {
      setExecutingAction(false);
    }
  }

  const onTogglingAllRecords = (newLockState: boolean): void => { 
    Promise.resolve().then(async () => {
      if (itemsWithMetadata === undefined) {
        return;
      }
      
      Log.info(LOG_SOURCE, `Toggling record status for all items`);
            
      setActionStatus(strings.CheckingItems);
      setItemsState([]);
      setNotification(undefined);
      setExecutingAction(true);

      // Trigger re-render table
      setItemsWithMetadata([...itemsWithMetadata]);

      try {
        const allItemsWithMetadata: IItemMetadata[] = await fetchAllItems();

        const itemsToToggle = allItemsWithMetadata.filter(i => !i.isFolder && i.isRecordTypeLabel && i.isRecordLocked !== newLockState).map(i => i.id);
        setActionStatus(format(strings.TogglingItems, itemsToToggle.length));
        const responses = await props.onToggling(itemsToToggle, newLockState);
        const errorCount = responses.filter(r => !r.success).length;
        
        if (errorCount > 0) {
          setNotification({ message: format(strings.ToggleErrorForMultipleItems, newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase(), errorCount, listItemIds.length), notificationType: MessageBarType.warning });
        }
        else {
          setNotification({ message: format(strings.RecordStatusToggled, newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase()), notificationType: MessageBarType.success });
        }                

        let newItemsState: IItemState[] = [];
        for(const itemResponse of responses) {
          newItemsState = [...newItemsState.filter(i => i.listItemId !== itemResponse.listItemId), { listItemId: itemResponse.listItemId, toggling: false, errorToggling: itemResponse.errorMessage, clearing: false, errorClearing: false }];
        }

        setItemsState(newItemsState);
        await fetchData(listItemIds, 1, true, false);
      } catch (error) {
        setNotification({ message: error.message, notificationType: MessageBarType.error });
        Log.error(LOG_SOURCE, new Error(error.message)); 

        setItemsState([]);

        // Trigger re-render table
        setItemsWithMetadata([...itemsWithMetadata]);
      }
      finally {
        setExecutingAction(false);
        setActionStatus("");
      }
    }).catch(error => console.log(error));
  }
  
  const onClearingLabel = async (item: IItemMetadata): Promise<void> => {
    if (itemsState === undefined || itemsWithMetadata === undefined) {
      return;
    }

    Log.info(LOG_SOURCE, `Clearing label for '${item.name}'`);

    setItemsState([...itemsState.filter(i => i.listItemId !== item.id), { listItemId: item.id, toggling: false, errorToggling: undefined, clearing: true, errorClearing: false }]);
    setNotification(undefined);    
    setExecutingAction(true);

    // Trigger re-render table
    setItemsWithMetadata([...itemsWithMetadata]);

    try {
      const responses = await props.onClearing([item.id]);
      const success = responses[0].success;

      if (!success) {
        setNotification({ message: strings.ClearErrorForSingleItem, notificationType: MessageBarType.error });
      }
      else {
        setNotification({ message: strings.LabelCleared, notificationType: MessageBarType.success });
      }
            
      setItemsState([...itemsState.filter(i => i.listItemId !== item.id), { listItemId: item.id, toggling: false, errorToggling: undefined, clearing: false, errorClearing: responses[0].success === false }]);
      setListItemIds(listItemIds.filter(i => i !== item.id));
      setItemsWithMetadata([...itemsWithMetadata.filter(i => i.id !== item.id)]);
      setFetchedItems([...fetchedItems.filter(i => i.id !== item.id)]);      
    }
    catch (error) {
      setNotification({ message: error.message, notificationType: MessageBarType.error });
      Log.error(LOG_SOURCE, new Error(error.message));      

      setItemsState([...itemsState.filter(i => i.listItemId !== item.id), { listItemId: item.id, toggling: false, errorToggling: undefined, clearing: false, errorClearing: false }]);

      // Trigger re-render table
      setItemsWithMetadata([...itemsWithMetadata]);
    }
    finally {
      setExecutingAction(false);
    }
  }

  const onClearingAllLabels = (): void => {    
    Promise.resolve().then(async () => {
      if (itemsWithMetadata === undefined) {
        return;
      }
      
      Log.info(LOG_SOURCE, `Clearing all labels`);
            
      setActionStatus(strings.CheckingItems);
      setItemsState([]);
      setNotification(undefined);
      setExecutingAction(true);

      // Trigger re-render table
      setItemsWithMetadata([...itemsWithMetadata]);

      try {
        const allItemsWithMetadata: IItemMetadata[] = await fetchAllItems();

        const itemsToClear = allItemsWithMetadata.filter(i => i.retentionLabel !== "").map(i => i.id);
        setActionStatus(format(strings.ClearingItems, itemsToClear.length));
        const responses = await props.onClearing(itemsToClear);
        const errorCount = responses.filter(r => !r.success).length;
        
        if (errorCount > 0) {
          setNotification({ message: format(strings.ClearErrorForMultipleItems, errorCount, listItemIds.length), notificationType: MessageBarType.warning });
        }
        else {
          setNotification({ message: strings.LabelCleared, notificationType: MessageBarType.success });
        }        
        
        let newItemsState: IItemState[] = [];
        let newListItemIds: number[] = listItemIds;
        for(const itemResponse of responses) {
          newItemsState = [...newItemsState.filter(i => i.listItemId !== itemResponse.listItemId), { listItemId: itemResponse.listItemId, toggling: false, errorToggling: undefined, clearing: false, errorClearing: !itemResponse.success }];
          if (itemResponse.success) {
            newListItemIds = newListItemIds.filter(i => i !== itemResponse.listItemId);
          }
        }

        setItemsState(newItemsState);
        setListItemIds([...newListItemIds]);
        await fetchData(newListItemIds, 1, true, false);
      } catch (error) {
        setNotification({ message: error.message, notificationType: MessageBarType.error });
        Log.error(LOG_SOURCE, new Error(error.message)); 

        // Trigger re-render table
        setItemsWithMetadata([...itemsWithMetadata]);
      }
      finally {
        setExecutingAction(false);
        setActionStatus("");
      }
    }).catch(error => console.log(error));
  }

  const onRenderItemColumn = (item: IItemMetadata, index: number, column: ICustomColumn): JSX.Element => {
    return <ItemColumn item={item} itemState={itemsState.filter(i => i.listItemId === item.id)[0]} column={column} onToggling={onTogglingRecord} onClearing={onClearingLabel} />;
  }
  
  const menuProps: IContextualMenuProps = {  
    items: [
      {
        key: 'lockRecords',
        text: strings.LockRecords,
        title: strings.LockRecordsTooltip,
        onClick: () => onTogglingAllRecords(true),
        iconProps: { iconName: 'Lock' },
      },
      {
        key: 'unlockRecords',
        text: strings.UnlockRecords,
        title: strings.UnlockRecordsTooltip,
        onClick: () => onTogglingAllRecords(false),
        iconProps: { iconName: 'Unlock' },
      },
      {
        key: 'clearAllLabels',
        text: strings.ClearLabels,
        title: strings.ClearLabelsTooltip,
        onClick: () => onClearingAllLabels(),
        iconProps: { iconName: 'Untag' },
      },
    ],
    directionalHintFixed: true,
  };

  const getMenu = (props: IContextualMenuProps): JSX.Element => {
    return <ContextualMenu {...props} />;
  }

  const getPage = (page: number): void => {
    Log.info(LOG_SOURCE, `Fetching page ${page}`);        
    fetchData(listItemIds, page).catch((error) => { console.log(error); });
  }

  useEffect(() => {
    fetchData(listItemIds, pageNumber).catch((error) => { console.log(error); });
  }, []);

  let paginationButtons: IButtonProps[] = [];

  if (totalPages > 1) {
    paginationButtons = [
      { iconProps: { iconName: "ChevronLeft"}, onClick: () => getPage(pageNumber-1), disabled: pageNumber === 1, title: pageNumber === 1 ? strings.IsFirstPage : format(strings.ToPage, pageNumber-1) }, 
      { iconProps: { iconName: "ChevronRight"}, onClick: () => getPage(pageNumber+1), disabled: pageNumber === totalPages, title: pageNumber === totalPages ? strings.IsLastPage : format(strings.ToPage, pageNumber+1) }
    ];
  }

  return <>
    <Dialog maxWidth={"1200px"} hidden={false} dialogContentProps={{ type:DialogType.largeHeader, title: strings.RetentionControlsHeader, responsiveMode: ResponsiveMode.small, topButtonsProps: paginationButtons, showCloseButton: true, onDismiss: props.onClose}}>
      {notification ? (
        <MessageBar styles={messageBarStyles} messageBarType={notification.notificationType}>
          {notification.message}
        </MessageBar>
      ) : (
        <></>
      )}    
      
      <ShimmeredDetailsList    
        items={itemsWithMetadata ?? []}
        columns={itemMetadataColumns}
        compact={true}   
        selectionMode={SelectionMode.none}
        onRenderItemColumn={onRenderItemColumn}          
        enableShimmer={loading}
        shimmerLines={shimmerLines}
      />
      {
        !loading && (!itemsWithMetadata || itemsWithMetadata?.length === 0) ?
          <MessageBar styles={messageBarStyles} messageBarType={MessageBarType.info}>
            {strings.NoLabelsApplied}
          </MessageBar>
        : <></>
      }      
      <DialogFooter styles={dialogFooterStyles}>            
        <DefaultButton onClick={props.onClose} text={strings.CloseModal} />
        <Stack horizontal tokens={{ childrenGap: 10 }}>

          <PrimaryButton menuProps={menuProps} menuAs={getMenu} iconProps={{ iconName: "MultiSelect"}} text={strings.TakeBulkActionsSelectedItems} title={strings.TakeBulkActionsSelectedItemsTooltip} disabled={!itemsWithMetadata || itemsWithMetadata?.length === 0 || executingAction} />
          { executingAction ? <Spinner label={actionStatus} labelPosition="right" size={SpinnerSize.small} /> : <></> }
        </Stack>
      </DialogFooter>
    </Dialog>    
  </>;
};