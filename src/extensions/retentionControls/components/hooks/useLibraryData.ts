import { useState, useCallback } from "react";
import { MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Log } from "@microsoft/sp-core-library";
import * as strings from "RetentionControlsCommandSetStrings";
import { LOG_SOURCE } from "../../RetentionControlsCommandSet";
import { IItemMetadata } from "../../../../shared/interfaces/IItemMetadata";
import { IDriveItem } from "../../../../shared/interfaces/IDriveItem";
import { IPagedDriveItems } from "../../../../shared/interfaces/IPagedDriveItems";
import { INotification } from "../../../../shared/interfaces/INotification";
import { flattenItemMetadata, flattenItemMetadataList, updateObjectProperties } from "../../../../shared/utils";
import { ILibraryDataHook } from "../types/LibraryViewTypes";

interface UseLibraryDataProps {
  onFetching: (listItemIds: number[]) => Promise<IDriveItem[]>;
  onFetchingPaged: (pageSize: number, nextLink?: string) => Promise<IPagedDriveItems>;
  fetchPageSize: number;
  pageSize: number;
}

export const useLibraryData = ({ onFetching, onFetchingPaged, fetchPageSize, pageSize }: UseLibraryDataProps): ILibraryDataHook => {
  const [nextLink, setNextLink] = useState<string | undefined>();
  const [loading, setLoading] = useState<boolean>(true);
  const [notification, setNotification] = useState<INotification | undefined>();
  const [totalPages, setTotalPages] = useState<number>(1);
  const [pageNumber, setPageNumber] = useState<number>(1);
  const [fetchedItems, setFetchedItems] = useState<IItemMetadata[]>([]);
  const [itemsWithMetadata, setItemsWithMetadata] = useState<IItemMetadata[] | undefined>();

  const refreshItemMetadata = useCallback(async (listItemId: number): Promise<void> => {
    if (itemsWithMetadata === undefined) {
      return;
    }

    try {
      const response = await onFetching([listItemId]);

      if (!response || response.length === 0) {
        throw new Error(strings.ErrorFetchingData);
      }

      const updatedItem = flattenItemMetadata(response[0]) as IItemMetadata;
      const fetchedItem = fetchedItems.filter(i => i.id === updatedItem.id)[0];
      const item = itemsWithMetadata.filter(i => i.id === updatedItem.id)[0];

      updateObjectProperties(fetchedItem, updatedItem);
      updateObjectProperties(item, updatedItem);

      setFetchedItems([...fetchedItems]);
      setItemsWithMetadata([...itemsWithMetadata]);
    } catch (error) {
      const message = strings.ErrorFetchingData + ": " + error.message;
      Log.error(LOG_SOURCE, new Error(message));
      setNotification({ message: message, notificationType: MessageBarType.error });
    }
  }, [itemsWithMetadata, fetchedItems, onFetching]);

  const fetchData = useCallback(async (page: number, showLoader: boolean = true, appendToFetchedList: boolean = true): Promise<void> => {
    try {
      if (showLoader) {
        setLoading(true);
      }

      let fetchedItemsList = fetchedItems;

      if (!appendToFetchedList) {
        fetchedItemsList = [];
      }

      setPageNumber(page);
      const retrieveNewItems = (page === 1 && fetchedItemsList.length === 0) || 
                              (page * pageSize > (fetchedItemsList?.length ?? 0) && nextLink !== undefined);

      if (retrieveNewItems) {
        const response = await onFetchingPaged(fetchPageSize, nextLink);
        const newItems = flattenItemMetadataList(response.items);

        for (const item of newItems) {
          if (!fetchedItemsList.some(i => i.driveItemId === item.driveItemId)) {
            fetchedItemsList.push(item);
          }
        }

        setNextLink(response.nextLink);
        setFetchedItems([...fetchedItemsList]);
      }

      const slice = fetchedItemsList.slice((page - 1) * pageSize, page * pageSize);
      setItemsWithMetadata(slice);
      setTotalPages(Math.ceil(fetchedItemsList.length / pageSize));

      if (showLoader) {
        setLoading(false);
      }
    } catch (error) {
      const message = strings.ErrorFetchingData + ": " + error.message;
      Log.error(LOG_SOURCE, new Error(message));
      setNotification({ message: message, notificationType: MessageBarType.error });

      if (showLoader) {
        setLoading(false);
      }
    }
  }, [fetchedItems, nextLink, pageSize, fetchPageSize, onFetchingPaged]);

  const clearNotification = useCallback(() => {
    setNotification(undefined);
  }, []);

  const updateItemsInPlace = useCallback(() => {
    if (itemsWithMetadata) {
      setItemsWithMetadata([...itemsWithMetadata]);
    }
  }, [itemsWithMetadata]);

  const removeItemFromLists = useCallback((itemId: number) => {
    if (itemsWithMetadata) {
      setItemsWithMetadata([...itemsWithMetadata.filter(i => i.id !== itemId)]);
    }
    setFetchedItems([...fetchedItems.filter(i => i.id !== itemId)]);
  }, [itemsWithMetadata, fetchedItems]);

  return {
    // State
    loading,
    notification,
    totalPages,
    pageNumber,
    itemsWithMetadata,
    
    // Actions
    fetchData,
    refreshItemMetadata,
    clearNotification,
    updateItemsInPlace,
    removeItemFromLists,
    setNotification
  };
};