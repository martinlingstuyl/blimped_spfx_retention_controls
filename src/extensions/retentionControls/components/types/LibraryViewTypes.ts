import { IItemMetadata } from "../../../../shared/interfaces/IItemMetadata";
import { IItemState } from "../../../../shared/interfaces/IItemState";
import { IBatchItemResponse } from "../../../../shared/interfaces/IBatchErrorResponse";
import { IDriveItem } from "../../../../shared/interfaces/IDriveItem";
import { IPagedDriveItems } from "../../../../shared/interfaces/IPagedDriveItems";
import { INotification } from "../../../../shared/interfaces/INotification";

// Props interface for the main LibraryView component
export interface ILibraryViewProps {
  isServedFromLocalhost: boolean;
  hasEditPermissions: boolean;
  onClose: () => void;
  onFetching: (listItemIds: number[]) => Promise<IDriveItem[]>;
  onFetchingPaged: (pageSize: number, nextLink?: string) => Promise<IPagedDriveItems>;
  onClearing: (listItemIds: number[]) => Promise<IBatchItemResponse[]>;
  onToggling: (listItemIds: number[], newLockstate: boolean) => Promise<IBatchItemResponse[]>;
}

// State interface for library data management
export interface ILibraryDataState {
  nextLink?: string;
  loading: boolean;
  notification?: INotification;
  totalPages: number;
  pageNumber: number;
  fetchedItems: IItemMetadata[];
  itemsWithMetadata?: IItemMetadata[];
}

// State interface for item actions
export interface IItemActionState {
  executingAction: boolean;
  actionStatus: string;
  itemsState: IItemState[];
}

// Action types for bulk operations
export type BulkActionType = 'toggle' | 'clear';

// Error context for better error handling
export interface IErrorContext {
  operation: string;
  itemName?: string;
  itemId?: number;
  additionalData?: Record<string, unknown>;
}

// Hook return types for better type safety
export interface ILibraryDataHook {
  loading: boolean;
  notification?: INotification;
  totalPages: number;
  pageNumber: number;
  itemsWithMetadata?: IItemMetadata[];
  fetchData: (page: number, showLoader?: boolean, appendToFetchedList?: boolean) => Promise<void>;
  refreshItemMetadata: (listItemId: number) => Promise<void>;
  clearNotification: () => void;
  updateItemsInPlace: () => void;
  removeItemFromLists: (itemId: number) => void;
  setNotification: (notification: INotification | undefined) => void;
}

export interface IItemActionsHook {
  executingAction: boolean;
  actionStatus: string;
  itemsState: IItemState[];
  onTogglingRecord: (item: IItemMetadata) => Promise<void>;
  onClearingLabel: (item: IItemMetadata) => Promise<void>;
  onTogglingAllRecords: (newLockState: boolean) => void;
  onClearingAllLabels: () => void;
}

export interface IPaginationHook {
  paginationButtons: import("@fluentui/react/lib/Button").IButtonProps[];
}

// Configuration type for better type safety
export interface ILibraryViewConfig {
  readonly FETCH_PAGE_SIZE: number;
  readonly PAGE_SIZE: number;
  readonly SHIMMER_LINES: number;
  readonly DIALOG_MAX_WIDTH: string;
  readonly BULK_ACTION_PAGE_SIZE: number;
}