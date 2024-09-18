import * as React from "react";
import { ICustomColumn } from "../../../shared/interfaces/ICustomColumn";
import { IItemMetadata } from "../../../shared/interfaces/IItemMetadata";
import { FontIcon } from "@fluentui/react/lib/Icon";
import { getBehaviorLabel } from "../../../shared/utils";
import { getFileTypeIconProps, FileIconType } from "@fluentui/react-file-type-icons";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import * as strings from "RetentionControlsCommandSetStrings";
import { classNames } from "../../../shared/styles";
import { format } from "@fluentui/react/lib/Utilities";
import { IItemState } from "../../../shared/interfaces/IItemState";
import ConfirmationDialogManager from "../ConfirmationDialogManager";
import AlertDialogManager from "../AlertDialogManager";

export interface IItemColumn {
  item: IItemMetadata;
  itemState?: IItemState;
  column: ICustomColumn;
  onClearing: (item: IItemMetadata) => void;
  onToggling: (item: IItemMetadata) => void;
}

export const ItemColumn: React.FC<IItemColumn> = (props) => {
  const { item, itemState, column, onToggling, onClearing } = props;

  const onToggleClick = async (item: IItemMetadata): Promise<void> => {
    if (!item.isRecordTypeLabel) {
      const dialog = new AlertDialogManager(strings.ToggleRecordForNonRecordLabelAlertTitle, strings.ToggleRecordForNonRecordLabelAlertMessage);
      await dialog.show();
    }
    else if (item.isFolder) {
      const dialog = new AlertDialogManager(strings.ToggleRecordForFolderAlertTitle, strings.ToggleRecordForFolderAlertMessage);
      await dialog.show();
    }
    else {
      onToggling(item)
    }
  }

  const onClearClick = async (item: IItemMetadata): Promise<void> => {
    if (item.isFolder) {
      const dialog = new ConfirmationDialogManager(strings.ClearLabelConfirmationTitle, strings.ClearLabelConfirmationMessage);
      dialog.onClosed((confirmed?: boolean) => {
        if (confirmed === true) {
          onClearing(item);
        }
      });
      await dialog.show();
    }
    else if (item.isRecordTypeLabel && !item.isRecordLocked) {
      const dialog = new AlertDialogManager(strings.CannotClearWhileUnlockedTitle, strings.CannotClearWhileUnlockedMessage);
      await dialog.show();
    }
    else {
      onClearing(item);
    }
  }

  if (column.key === "icon") {
    if (item.contentTypeId.indexOf("0x0120D520") > -1) {
      return <FontIcon {...getFileTypeIconProps({ type: FileIconType.docset, size: 16 })} />;
    } else if (item.contentTypeId.indexOf("0x0120") > -1) {
      return <FontIcon {...getFileTypeIconProps({ type: FileIconType.folder, size: 16 })} />;
    }

    const extension = item.name.substring(item.name.lastIndexOf("."));
    return <FontIcon {...getFileTypeIconProps({ extension: extension, size: 16 })} />;
  } else if (column.key === "name") {      
    return <span title={item.path}>{item.name}</span>;
  } else if (column.key === "behaviorDuringRetentionPeriod") {
    const fieldValue = getBehaviorLabel(item.behaviorDuringRetentionPeriod);
    return <span title={fieldValue}>{fieldValue}</span>;
  } else if (column.key === "isDeleteAllowed" || column.key === "isMetadataUpdateAllowed" || column.key === "isContentUpdateAllowed" || column.key === "isLabelUpdateAllowed") {
    const boolValue = item[column.key as keyof IItemMetadata];
    if (boolValue === true) {
      return <FontIcon iconName="Accept" className={classNames.green} title={column.title + ": " + strings.ToggleOnText} />;
    } else if (boolValue === false) {
      return <FontIcon className={classNames.red} iconName="Cancel" title={column.title + ": " + strings.ToggleOffText} />;
    } 

    return <></>;      
  } else if (column.key === "isRecordLocked") {
    if (itemState?.toggling) {
      return <Spinner size={SpinnerSize.xSmall} />;
    }    
    if (itemState?.errorToggling !== undefined) {      
      return <FontIcon iconName="Warning" className={classNames.red} title={format(strings.ToggleWarning, item.isRecordLocked === true ? strings.Locked : strings.Unlocked, itemState.errorToggling)} onClick={() => onToggleClick(item)} />;
    }
    if (item.isRecordLocked === true) {
      return <FontIcon iconName="LockSolid" className={classNames.dark} title={column.title + ": " + strings.Locked} onClick={() => onToggleClick(item)} />;
    } else if (item.isRecordLocked === false) {
      return <FontIcon iconName="Unlock" className={classNames.dark} title={column.title + ": " + strings.Unlocked} onClick={() => onToggleClick(item)} />;
    } 

    return <></>;      
  }
  else if (column.key === "clearLabel") {
    if (itemState?.clearing) {
      return <Spinner size={SpinnerSize.xSmall} />;
    }
    if (itemState?.errorClearing === true) {
      return <FontIcon iconName="Warning" className={classNames.red} title={strings.ClearLabelWarningTooltip} onClick={() => onClearClick(item)} />;
    }

    return <FontIcon iconName="Untag" className={classNames.dark} title={column.title} onClick={() => onClearClick(item)} />;
  }

  const fieldValue = item[column.key as keyof IItemMetadata] as string;
  return <span title={fieldValue}>{fieldValue}</span>;
};
