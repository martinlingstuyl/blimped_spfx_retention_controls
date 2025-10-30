import { useState, useCallback } from "react";
import { MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Log } from "@microsoft/sp-core-library";
import { format } from "@fluentui/react/lib/Utilities";
import * as strings from "RetentionControlsCommandSetStrings";
import { LOG_SOURCE } from "../../RetentionControlsCommandSet";
import { IItemMetadata } from "../../../../shared/interfaces/IItemMetadata";
import { IItemState } from "../../../../shared/interfaces/IItemState";
import { IBatchItemResponse } from "../../../../shared/interfaces/IBatchErrorResponse";
import { IPagedDriveItems } from "../../../../shared/interfaces/IPagedDriveItems";
import { flattenItemMetadataList } from "../../../../shared/utils";
import { IItemActionsHook } from "../types/LibraryViewTypes";

interface UseItemActionsProps {
  onToggling: (listItemIds: number[], newLockstate: boolean) => Promise<IBatchItemResponse[]>;
  onClearing: (listItemIds: number[]) => Promise<IBatchItemResponse[]>;
  onFetchingPaged: (pageSize: number, nextLink?: string) => Promise<IPagedDriveItems>;
  refreshItemMetadata: (listItemId: number) => Promise<void>;
  updateItemsInPlace: () => void;
  removeItemFromLists: (itemId: number) => void;
  setNotification: (notification: unknown) => void;
  clearNotification: () => void;
}

export const useItemActions = ({
  onToggling,
  onClearing,
  onFetchingPaged,
  refreshItemMetadata,
  updateItemsInPlace,
  removeItemFromLists,
  setNotification,
  clearNotification
}: UseItemActionsProps): IItemActionsHook => {
  const [executingAction, setExecutingAction] = useState<boolean>(false);
  const [actionStatus, setActionStatus] = useState<string>("");
  const [itemsState, setItemsState] = useState<IItemState[]>([]);

  const updateItemState = useCallback((itemId: number, state: Partial<IItemState>) => {
    setItemsState(prev => [
      ...prev.filter(i => i.listItemId !== itemId),
      { 
        listItemId: itemId, 
        toggling: false, 
        errorToggling: undefined, 
        clearing: false, 
        errorClearing: false,
        ...state 
      }
    ]);
  }, []);

  const onTogglingRecord = useCallback(async (item: IItemMetadata): Promise<void> => {
    Log.info(LOG_SOURCE, `Toggling record status for '${item.name}'`);

    updateItemState(item.id, { toggling: true });
    clearNotification();
    setExecutingAction(true);
    updateItemsInPlace();

    try {
      const newLockState = !item.isRecordLocked;
      const response = await onToggling([item.id], newLockState);
      const success = response[0].success;

      if (!success) {
        setNotification({
          message: format(strings.ToggleErrorForSingleItem, 
            newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase()) + 
            " " + response[0].errorMessage,
          notificationType: MessageBarType.error
        });
      } else {
        setNotification({
          message: format(strings.RecordStatusToggled, 
            newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase()),
          notificationType: MessageBarType.success
        });
      }

      updateItemState(item.id, {
        toggling: false,
        errorToggling: response[0].success === false ? response[0].errorMessage : undefined
      });
      
      await refreshItemMetadata(item.id);
    } catch (error) {
      setNotification({ message: error.message, notificationType: MessageBarType.error });
      Log.error(LOG_SOURCE, new Error(error.message));
      updateItemState(item.id, { toggling: true });
      updateItemsInPlace();
    } finally {
      setExecutingAction(false);
    }
  }, [onToggling, updateItemState, clearNotification, updateItemsInPlace, refreshItemMetadata, setNotification]);

  const onClearingLabel = useCallback(async (item: IItemMetadata): Promise<void> => {
    Log.info(LOG_SOURCE, `Clearing label for '${item.name}'`);

    updateItemState(item.id, { clearing: true });
    clearNotification();
    setExecutingAction(true);
    updateItemsInPlace();

    try {
      const responses = await onClearing([item.id]);
      const success = responses[0].success;

      if (!success) {
        setNotification({ message: strings.ClearErrorForSingleItem, notificationType: MessageBarType.error });
      } else {
        setNotification({ message: strings.LabelCleared, notificationType: MessageBarType.success });
      }

      updateItemState(item.id, {
        clearing: false,
        errorClearing: responses[0].success === false
      });
      
      removeItemFromLists(item.id);
    } catch (error) {
      setNotification({ message: error.message, notificationType: MessageBarType.error });
      Log.error(LOG_SOURCE, new Error(error.message));
      updateItemState(item.id, { clearing: false });
      updateItemsInPlace();
    } finally {
      setExecutingAction(false);
    }
  }, [onClearing, updateItemState, clearNotification, updateItemsInPlace, removeItemFromLists, setNotification]);

  const processBulkAction = useCallback(async (
    actionType: 'toggle' | 'clear',
    newLockState?: boolean
  ): Promise<void> => {
    const isToggleAction = actionType === 'toggle';
    Log.info(LOG_SOURCE, isToggleAction ? `Toggling record status for all items` : `Clearing all labels`);

    setActionStatus(`0 ${strings.ItemsDone}`);
    setItemsState([]);
    clearNotification();
    setExecutingAction(true);
    updateItemsInPlace();

    try {
      let newItemsState: IItemState[] = [];
      let errorCount = 0;
      let more = true;
      let nextLink = undefined;

      while (more) {
        const itemsPage: IPagedDriveItems = await onFetchingPaged(100, nextLink);
        
        let responses: IBatchItemResponse[];
        if (isToggleAction && newLockState !== undefined) {
          const itemsToToggle = flattenItemMetadataList(itemsPage.items)
            .filter(i => !i.isFolder && i.isRecordTypeLabel && i.isRecordLocked !== newLockState)
            .map(i => i.id);
          responses = await onToggling(itemsToToggle, newLockState);
        } else {
          responses = await onClearing(itemsPage.items.map(i => parseFloat(i.listItem.id)));
        }

        more = itemsPage.nextLink !== undefined && itemsPage.items.length > 0;
        nextLink = itemsPage.nextLink;
        errorCount += responses.filter(r => !r.success).length;

        for (const itemResponse of responses) {
          newItemsState = [
            ...newItemsState.filter(i => i.listItemId !== itemResponse.listItemId),
            {
              listItemId: itemResponse.listItemId,
              toggling: false,
              errorToggling: isToggleAction ? itemResponse.errorMessage : undefined,
              clearing: false,
              errorClearing: !isToggleAction ? itemResponse.success === false : false
            }
          ];
        }

        setActionStatus(`${newItemsState.length} ${strings.ItemsDone}`);
      }

      // Set appropriate notification based on results
      if (errorCount > 0) {
        const message = isToggleAction
          ? format(strings.ToggleErrorForMultipleItems, 
              newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase(), 
              errorCount, newItemsState.length)
          : format(strings.ClearErrorForMultipleItems, errorCount, newItemsState.length);
        setNotification({ message, notificationType: MessageBarType.warning });
      } else {
        const message = isToggleAction
          ? format(strings.RecordStatusToggledEntireLibrary, 
              newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase())
          : strings.LabelClearedForLibrary;
        setNotification({ message, notificationType: MessageBarType.success });
      }

      setItemsState(newItemsState);
    } catch (error) {
      setNotification({ message: error.message, notificationType: MessageBarType.error });
      Log.error(LOG_SOURCE, new Error(error.message));
      setItemsState([]);
      updateItemsInPlace();
    } finally {
      setExecutingAction(false);
      setActionStatus("");
    }
  }, [onToggling, onClearing, onFetchingPaged, clearNotification, updateItemsInPlace, setNotification]);

  const onTogglingAllRecords = useCallback((newLockState: boolean): void => {
    processBulkAction('toggle', newLockState).catch(error => console.log(error));
  }, [processBulkAction]);

  const onClearingAllLabels = useCallback((): void => {
    processBulkAction('clear').catch(error => console.log(error));
  }, [processBulkAction]);

  return {
    // State
    executingAction,
    actionStatus,
    itemsState,
    
    // Actions
    onTogglingRecord,
    onClearingLabel,
    onTogglingAllRecords,
    onClearingAllLabels
  };
};