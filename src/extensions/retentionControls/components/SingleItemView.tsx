import * as React from "react";
import { IDriveItem } from "../../../shared/interfaces/IDriveItem";
import * as strings from "RetentionControlsCommandSetStrings";
import { classNames, dialogFooterStyles, messageBarStyles, stackItemStyles, stackTokens } from "../../../shared/styles";
import { Stack } from "@fluentui/react/lib/Stack";
import { Shimmer } from "@fluentui/react/lib/Shimmer";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Link } from "@fluentui/react/lib/Link";
import { FontIcon } from "@fluentui/react/lib/Icon";
import { flattenItemMetadata, getBehaviorLabel } from "../../../shared/utils";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import Dialog, { DialogFooter, DialogType } from "@fluentui/react/lib/Dialog";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { useEffect, useState } from "react";
import { Log } from "@microsoft/sp-core-library";
import { LOG_SOURCE } from "../RetentionControlsCommandSet";
import { IItemState } from "../../../shared/interfaces/IItemState";
import { RowAccessor } from "@microsoft/sp-listview-extensibility";
import { IBatchItemResponse } from "../../../shared/interfaces/IBatchErrorResponse";
import { format } from "@fluentui/react/lib/Utilities";
import { ResponsiveMode } from "@fluentui/react/lib/ResponsiveMode";
import { INotification } from "../../../shared/interfaces/INotification";
import AlertDialogManager from "../AlertDialogManager";
import ConfirmationDialogManager from "../ConfirmationDialogManager";
import { IItemMetadata } from "../../../shared/interfaces/IItemMetadata";
initializeIcons();

export interface ISingleItemView {
  listItems: readonly RowAccessor[];
  onClose: () => void;
  onFetching: (listItemIds: number[]) => Promise<IDriveItem[]>;
  onClearing: (listItemIds: number[]) => Promise<IBatchItemResponse[]>;
  onToggling: (listItemIds: number[], newLockState: boolean) => Promise<IBatchItemResponse[]>;
}

export const SingleItemView: React.FC<ISingleItemView> = (props) => {
  const { listItems } = props;
  const listItem = listItems[0];
  const listItemId = listItem !== undefined ? parseFloat(listItem.getValueByName("ID")) : undefined;
  const fileName = listItem?.getValueByName("FileLeafRef");
  const initialRetentionLabel = listItem?.getValueByName("_ComplianceTag");

  const [loading, setLoading] = useState<boolean>(true);
  const [notification, setNotification] = useState<INotification | undefined>();
  const [itemDetails, setItemDetails] = useState<IItemMetadata | undefined>();
  const [itemState, setItemState] = useState<Partial<IItemState>>({ clearing: false, toggling: false, errorClearing: false });

  const labelAppliedDate = itemDetails?.labelAppliedDate ? itemDetails?.labelAppliedDate : "N/A";
  const eventDate = itemDetails?.eventDate !== undefined && itemDetails?.eventDate?.indexOf("9999") === -1 ? new Date(itemDetails?.eventDate).toLocaleDateString() : undefined;
  
  const fetchData = async (): Promise<void> => {
    try {
      setLoading(true);

      if (listItemId) {
        const response = await props.onFetching([listItemId]);    
        setItemDetails(flattenItemMetadata(response[0]));
      }
      else
        setNotification({ message: strings.NoLabelApplied, notificationType: MessageBarType.info });      

      setLoading(false);
    } catch (error) {
      const message = strings.ErrorFetchingData + ": " + error.message;
      Log.error(LOG_SOURCE, new Error(message));
      setNotification({ message: message, notificationType: MessageBarType.error });
      setLoading(false);
    }
  };
  
  const onClearingLabel = async (): Promise<void> => {
    if (!itemDetails || !listItemId) {
      return;
    }

    Log.info(LOG_SOURCE, `Clearing label for '${itemDetails.name}'`);

    setNotification(undefined);
    setItemState({ ...itemState, clearing: true, errorClearing: false });

    try {
      const responses = await props.onClearing([itemDetails.id]);
      const isError = responses.every((r) => r.success === false);
      const newItemDetails = await props.onFetching([listItemId]);    
      setItemDetails(flattenItemMetadata(newItemDetails[0]));

      if (isError) {
        setNotification({ message: strings.ClearErrorForSingleItem, notificationType: MessageBarType.error });
      }              
      else if (!isError && newItemDetails.every(d => d.retentionLabel?.name === undefined)) {
        setNotification({ message: strings.LabelCleared, notificationType: MessageBarType.success });
      }

      setItemState({ ...itemState, clearing: false, errorClearing: isError === true });
    }
    catch (error) {      
      setNotification({ message: error.message, notificationType: MessageBarType.error });
      Log.error(LOG_SOURCE, new Error(error.message));
      setItemState({ ...itemState, clearing: false, errorClearing: true });
    }    
  };

  const onTogglingLabel = async (): Promise<void> => {
    if (!itemDetails || !listItemId) {
      return;
    }

    Log.info(LOG_SOURCE, `Toggling record state for '${itemDetails.name}'`);

    setNotification(undefined);
    setItemState({ ...itemState, toggling: true, errorToggling: undefined });

    try {
      const newLockState = itemDetails.isRecordLocked === true ? false : true;
      const responses = await props.onToggling([itemDetails.id], newLockState);
      const isError = responses.every((r) => r.success === false);
      const newItemDetails = await props.onFetching([listItemId]);    
      setItemDetails(flattenItemMetadata(newItemDetails[0]));

      if (isError) {
        setNotification({ message: strings.ClearErrorForSingleItem, notificationType: MessageBarType.error });
      }              
      else if (!isError) {
        setNotification({ message: format(strings.RecordStatusToggled, newLockState === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase()), notificationType: MessageBarType.success });
      }

      setItemState({ ...itemState, toggling: false, errorToggling: responses[0].errorMessage });
    }
    catch (error) {
      setNotification({ message: error.message, notificationType: MessageBarType.error });
      Log.error(LOG_SOURCE, new Error(error.message));
      setItemState({ ...itemState, toggling: false });
    }    
  };

  const onToggleClick = async (): Promise<void> => {
    if (!itemDetails || !listItemId) {
      return;
    }

    if (itemDetails.isFolder) {
      const dialog = new AlertDialogManager(strings.ToggleRecordForFolderAlertTitle, strings.ToggleRecordForFolderAlertMessage);
      await dialog.show();
    }
    else if (!itemDetails.isRecordTypeLabel) {
      const dialog = new AlertDialogManager(strings.ToggleRecordForNonRecordLabelAlertTitle, strings.ToggleRecordForNonRecordLabelAlertMessage);
      await dialog.show();
    }
    else {
      await onTogglingLabel();
    }
  }

  const onClearClick = async (): Promise<void> => {
    if (!itemDetails || !listItemId) {
      return;
    }

    if (itemDetails.isFolder) {
      const dialog = new ConfirmationDialogManager(strings.ClearLabelConfirmationTitle, strings.ClearLabelConfirmationMessage);
      dialog.onClosed(async (confirmed?: boolean) => {
        if (confirmed === true) {
          await onClearingLabel();
        }
      });
      await dialog.show();
    }
    else if (itemDetails.isRecordTypeLabel && !itemDetails.isRecordLocked) {
      const dialog = new AlertDialogManager(strings.CannotClearWhileUnlockedTitle, strings.CannotClearWhileUnlockedMessage);
      await dialog.show();
    }
    else {
      await onClearingLabel();
    }
  }

  useEffect(() => {
    fetchData().catch((error) => { console.log(error); });
  }, []);

  return <>
    <Dialog maxWidth={"600px"} hidden={false} dialogContentProps={{ type:DialogType.largeHeader, title: strings.RetentionControlsHeader, responsiveMode: ResponsiveMode.small, showCloseButton: true, onDismiss: props.onClose}}>
      {notification ? (
        <MessageBar styles={messageBarStyles} messageBarType={notification.notificationType}>
          {notification.message}
        </MessageBar>
      ) : (
        <></>
      )}    
      {
        listItem ? <>
          <Stack tokens={{ childrenGap: 10 }}>
            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <strong>{strings.FileName}</strong>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <>
                  {fileName}
                </>
              </Stack.Item>
            </Stack>
            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <strong>{strings.RetentionLabelApplied}</strong>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <>
                  {
                    loading ? initialRetentionLabel : itemDetails?.retentionLabel ?? strings.None
                  }
                  {
                    !loading && itemDetails?.retentionLabel !== undefined ? 
                      <Link disabled={loading || itemState.clearing} onClick={() => onClearClick()} style={{ marginLeft: "10px" }}>
                        {itemState.clearing ? <Spinner size={SpinnerSize.xSmall} style={{ marginRight: "10px" }} labelPosition="right" label={strings.Clearing} /> : <>{strings.ClearLabel}</>}
                      </Link>
                    : <></>
                  }                  
                </>
              </Stack.Item>
            </Stack>
            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={250} height={16} isDataLoaded={!loading}>
                  <strong>{strings.RetentionLabelApplicationDate}</strong>
                </Shimmer>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={250} height={16} isDataLoaded={!loading}>
                  <div>{labelAppliedDate}</div>
                </Shimmer>
              </Stack.Item>
            </Stack>
            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={250} height={16} isDataLoaded={!loading}>
                  <strong>{strings.RetentionLabelAppliedBy}</strong>
                </Shimmer>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={250} height={16} isDataLoaded={!loading}>
                  <div>{itemDetails?.labelAppliedBy}</div>
                </Shimmer>
              </Stack.Item>
            </Stack>

            {eventDate ? (
              <Stack horizontal tokens={stackTokens}>
                <Stack.Item grow={1} styles={stackItemStyles}>
                  <Shimmer width={250} height={16} isDataLoaded={!loading}>
                    <strong>{strings.RetentionLabelEventDate}</strong>
                  </Shimmer>
                </Stack.Item>
                <Stack.Item grow={1} styles={stackItemStyles}>
                  <Shimmer width={250} height={16} isDataLoaded={!loading}>
                    <div>{eventDate}</div>
                  </Shimmer>
                </Stack.Item>
              </Stack>
            ) : (
              <></>
            )}

            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={220} height={16} isDataLoaded={!loading}>
                  <strong>{strings.BehaviorDuringRetentionPeriod}</strong>
                </Shimmer>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={150} height={16} isDataLoaded={!loading}>
                  {getBehaviorLabel(itemDetails?.behaviorDuringRetentionPeriod)}
                </Shimmer>
              </Stack.Item>
            </Stack>

            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={170} height={16} isDataLoaded={!loading}>
                  <strong>{strings.IsMetadataUpdateAllowed}</strong>
                  <FontIcon style={{ marginLeft: "6px", cursor: "pointer" }} iconName="Info" title={strings.IsMetadataUpdateAllowedTooltip} className={classNames.blue} />
                </Shimmer>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={50} height={16} isDataLoaded={!loading}>
                  {itemDetails?.isMetadataUpdateAllowed === false ? (
                    <>
                      <FontIcon className={classNames.red} iconName="Cancel" /> {strings.ToggleOffText}
                    </>
                  ) : (
                    <>
                      <FontIcon iconName="Accept" className={classNames.green} /> {strings.ToggleOnText}
                    </>
                  )}
                </Shimmer>
              </Stack.Item>
            </Stack>
            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={160} height={16} isDataLoaded={!loading}>
                  <strong>{strings.IsContentUpdateAllowed}</strong>
                  <FontIcon style={{ marginLeft: "6px", cursor: "pointer" }} iconName="Info" title={strings.IsContentUpdateAllowedTooltip} className={classNames.blue} />
                </Shimmer>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={50} height={16} isDataLoaded={!loading}>
                  {itemDetails?.isContentUpdateAllowed === false ? (
                    <>
                      <FontIcon className={classNames.red} iconName="Cancel" /> {strings.ToggleOffText}
                    </>
                  ) : (
                    <>
                      <FontIcon iconName="Accept" className={classNames.green} /> {strings.ToggleOnText}
                    </>
                  )}
                </Shimmer>
              </Stack.Item>
            </Stack>
            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={100} height={16} isDataLoaded={!loading}>
                  <strong>{strings.IsDeleteAllowed}</strong>
                  <FontIcon style={{ marginLeft: "6px", cursor: "pointer" }} iconName="Info" title={strings.IsDeleteAllowedTooltip} className={classNames.blue} />
                </Shimmer>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={50} height={16} isDataLoaded={!loading}>
                  {itemDetails?.isDeleteAllowed === false ? (
                    <>
                      <FontIcon className={classNames.red} iconName="Cancel" /> {strings.ToggleOffText}
                    </>
                  ) : (
                    <>
                      <FontIcon iconName="Accept" className={classNames.green} /> {strings.ToggleOnText}
                    </>
                  )}
                </Shimmer>
              </Stack.Item>
            </Stack>
            <Stack horizontal tokens={stackTokens}>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={140} height={16} isDataLoaded={!loading}>
                  <strong>{strings.IsLabelUpdateAllowed}</strong>
                  <FontIcon style={{ marginLeft: "6px", cursor: "pointer" }} iconName="Info" title={strings.IsLabelUpdateAllowedTooltip} className={classNames.blue} />
                </Shimmer>
              </Stack.Item>
              <Stack.Item grow={1} styles={stackItemStyles}>
                <Shimmer width={50} height={16} isDataLoaded={!loading}>
                  {itemDetails?.isLabelUpdateAllowed === false ? (
                    <>
                      <FontIcon className={classNames.red} iconName="Cancel" /> {strings.ToggleOffText}
                    </>
                  ) : (
                    <>
                      <FontIcon iconName="Accept" className={classNames.green} /> {strings.ToggleOnText}
                    </>
                  )}
                </Shimmer>
              </Stack.Item>
            </Stack>

            {itemDetails?.behaviorDuringRetentionPeriod === "retainAsRecord" ? (
              <>
                <Stack horizontal tokens={stackTokens}>
                  <Stack.Item grow={1} styles={stackItemStyles}>
                    <Shimmer width={100} height={16} isDataLoaded={!loading}>
                      <strong>{strings.RecordStatus}</strong>
                      <FontIcon style={{ marginLeft: "6px", cursor: "pointer" }} iconName="Info" title={strings.RecordStatusTooltip} className={classNames.blue} />
                    </Shimmer>
                  </Stack.Item>
                  <Stack.Item grow={1} styles={stackItemStyles}>
                    <Shimmer width={50} height={16} isDataLoaded={!loading}>
                      <>
                        {itemDetails.isRecordLocked === true ? (
                          <>
                            <FontIcon iconName="LockSolid" /> {strings.Locked}
                          </>
                        ) : (
                          <>
                            <FontIcon iconName="Unlock" /> {strings.Unlocked}
                          </>
                        )}
                        <Link disabled={itemState.toggling} onClick={() => onToggleClick()} style={{ marginLeft: "10px" }}>
                          {itemState.toggling ? <Spinner size={SpinnerSize.xSmall} style={{ marginRight: "10px" }} labelPosition="right" label={strings.Toggling} /> : <>{strings.ToggleLockStatus}</>}
                        </Link>
                      </>
                    </Shimmer>
                  </Stack.Item>
                </Stack>
              </>
            ) : (
              <></>
            )}
          </Stack>
        </> : <></>
      }
      <DialogFooter styles={dialogFooterStyles}>
        <DefaultButton onClick={props.onClose} text={strings.CloseModal} />
      </DialogFooter>
    </Dialog>
  </>;
};
