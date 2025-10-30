import * as React from "react";
import { useEffect } from "react";
import * as strings from "RetentionControlsCommandSetStrings";
import { SelectionMode } from "@fluentui/react/lib/Utilities";
import { ShimmeredDetailsList } from "@fluentui/react/lib/ShimmeredDetailsList";
import { IItemMetadata } from "../../../shared/interfaces/IItemMetadata";
import { ICustomColumn } from "../../../shared/interfaces/ICustomColumn";
import { ItemColumn } from "./ItemColumn";
import Dialog, { DialogFooter, DialogType } from "@fluentui/react/lib/Dialog";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { dialogFooterStyles, messageBarStyles } from "../../../shared/styles";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { ResponsiveMode } from "@fluentui/react/lib/ResponsiveMode";
import { itemMetadataColumns } from "../../../shared/constants";
import { Log } from "@microsoft/sp-core-library";
import { LOG_SOURCE } from "../RetentionControlsCommandSet";
import { useLibraryData } from "./hooks/useLibraryData";
import { useItemActions } from "./hooks/useItemActions";
import { usePagination } from "./hooks/usePagination";
import { NotificationBar } from "./ui/NotificationBar";
import { BulkActionsMenu } from "./ui/BulkActionsMenu";
import { LIBRARY_VIEW_CONFIG } from "./constants";
import { ILibraryViewProps } from "./types/LibraryViewTypes";
import { errorHandler } from "./utils/ErrorHandler";

export const LibraryView: React.FC<ILibraryViewProps> = (props) => {
  // Initialize error handler
  React.useEffect(() => {
    return () => {
      errorHandler.clearNotification();
    };
  }, []);

  const libraryData = useLibraryData({
    onFetching: props.onFetching,
    onFetchingPaged: props.onFetchingPaged,
    fetchPageSize: LIBRARY_VIEW_CONFIG.FETCH_PAGE_SIZE,
    pageSize: LIBRARY_VIEW_CONFIG.PAGE_SIZE
  });

  // Set up error handler callback
  React.useEffect(() => {
    errorHandler.setNotificationCallback(libraryData.setNotification);
  }, [libraryData.setNotification]);

  const itemActions = useItemActions({
    onToggling: props.onToggling,
    onClearing: props.onClearing,
    onFetchingPaged: props.onFetchingPaged,
    refreshItemMetadata: libraryData.refreshItemMetadata,
    updateItemsInPlace: libraryData.updateItemsInPlace,
    removeItemFromLists: libraryData.removeItemFromLists,
    setNotification: libraryData.setNotification,
    clearNotification: libraryData.clearNotification
  });

  const handlePageChange = React.useCallback((page: number) => {
    Log.info(LOG_SOURCE, `Fetching page ${page}`);
    libraryData.fetchData(page).catch((error) => {
      console.log(error);
    });
  }, [libraryData.fetchData]);

  const { paginationButtons } = usePagination({
    totalPages: libraryData.totalPages,
    pageNumber: libraryData.pageNumber,
    onPageChange: handlePageChange
  });

  const onRenderItemColumn = React.useCallback((
    item: IItemMetadata, 
    index: number, 
    column: ICustomColumn
  ): JSX.Element => {
    const itemState = itemActions.itemsState.filter(i => i.listItemId === item.id)[0];
    return (
      <ItemColumn
        item={item}
        itemState={itemState}
        column={column}
        hasEditPermissions={props.hasEditPermissions}
        onToggling={itemActions.onTogglingRecord}
        onClearing={itemActions.onClearingLabel}
      />
    );
  }, [itemActions.itemsState, itemActions.onTogglingRecord, itemActions.onClearingLabel, props.hasEditPermissions]);

  // Initialize data on component mount
  useEffect(() => {
    libraryData.fetchData(libraryData.pageNumber).catch((error) => {
      console.log(error);
    });
  }, []);

  const hasItems = Boolean(libraryData.itemsWithMetadata?.length);
  const showEmptyMessage = !libraryData.loading && !hasItems;

  return (
    <Dialog
      maxWidth={LIBRARY_VIEW_CONFIG.DIALOG_MAX_WIDTH}
      hidden={false}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: strings.RetentionControlsHeader,
        responsiveMode: ResponsiveMode.small,
        topButtonsProps: paginationButtons,
        showCloseButton: true,
        onDismiss: props.onClose
      }}
    >
      <NotificationBar
        notification={libraryData.notification}
        isServedFromLocalhost={props.isServedFromLocalhost}
      />

      <ShimmeredDetailsList
        items={libraryData.itemsWithMetadata ?? []}
        columns={itemMetadataColumns}
        compact={true}
        selectionMode={SelectionMode.none}
        onRenderItemColumn={onRenderItemColumn}
        enableShimmer={libraryData.loading}
        shimmerLines={LIBRARY_VIEW_CONFIG.SHIMMER_LINES}
      />

      {showEmptyMessage && (
        <MessageBar 
          styles={messageBarStyles} 
          messageBarType={MessageBarType.info}
        >
          {strings.NoLabelsAppliedEntireLibrary}
        </MessageBar>
      )}

      <DialogFooter styles={dialogFooterStyles}>
        <DefaultButton 
          onClick={props.onClose} 
          text={strings.CloseModal} 
        />
        
        <BulkActionsMenu
          hasEditPermissions={props.hasEditPermissions}
          hasItems={hasItems}
          executingAction={itemActions.executingAction}
          actionStatus={itemActions.actionStatus}
          onToggleAllRecords={itemActions.onTogglingAllRecords}
          onClearAllLabels={itemActions.onClearingAllLabels}
        />
      </DialogFooter>
    </Dialog>
  );
};