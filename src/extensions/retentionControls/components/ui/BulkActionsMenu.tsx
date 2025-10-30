import * as React from "react";
import { ContextualMenu, IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { Stack } from "@fluentui/react/lib/Stack";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import * as strings from "RetentionControlsCommandSetStrings";
import ConfirmationDialogManager from "../../ConfirmationDialogManager";

interface BulkActionsMenuProps {
  hasEditPermissions: boolean;
  hasItems: boolean;
  executingAction: boolean;
  actionStatus: string;
  onToggleAllRecords: (newLockState: boolean) => void;
  onClearAllLabels: () => void;
}

export const BulkActionsMenu: React.FC<BulkActionsMenuProps> = ({
  hasEditPermissions,
  hasItems,
  executingAction,
  actionStatus,
  onToggleAllRecords,
  onClearAllLabels
}) => {
  const showConfirmationDialog = (title: string, message: string, callback: () => void): void => {
    Promise.resolve().then(async () => {
      const dialog = new ConfirmationDialogManager(title, message);
      dialog.onClosed(async (confirmed?: boolean) => {
        if (confirmed === true) {
          callback();
        }
      });
      await dialog.show();
    }).catch(error => console.log(error));
  };

  const menuProps: IContextualMenuProps = {
    items: [
      {
        key: 'lockRecords',
        text: strings.LockRecords,
        title: strings.LockRecordsTooltip,
        onClick: () => showConfirmationDialog(
          strings.ConfirmEntireLibraryImpactTitle,
          strings.ConfirmEntireLibraryImpactMessage,
          () => onToggleAllRecords(true)
        ),
        iconProps: { iconName: 'Lock' },
        disabled: !hasEditPermissions,
      },
      {
        key: 'unlockRecords',
        text: strings.UnlockRecords,
        title: strings.UnlockRecordsTooltip,
        onClick: () => showConfirmationDialog(
          strings.ConfirmEntireLibraryImpactTitle,
          strings.ConfirmEntireLibraryImpactMessage,
          () => onToggleAllRecords(false)
        ),
        iconProps: { iconName: 'Unlock' },
        disabled: !hasEditPermissions,
      },
      {
        key: 'clearAllLabels',
        text: strings.ClearLabels,
        title: strings.ClearLabelsTooltip,
        onClick: () => showConfirmationDialog(
          strings.ConfirmEntireLibraryImpactTitle,
          strings.ConfirmEntireLibraryImpactMessage,
          () => onClearAllLabels()
        ),
        iconProps: { iconName: 'Untag' },
        disabled: !hasEditPermissions,
      },
    ],
    directionalHintFixed: true,
  };

  const getMenu = (props: IContextualMenuProps): JSX.Element => {
    return <ContextualMenu {...props} />;
  };

  return (
    <Stack horizontal tokens={{ childrenGap: 10 }}>
      {executingAction && (
        <Spinner 
          label={actionStatus} 
          labelPosition="right" 
          size={SpinnerSize.small} 
        />
      )}
      <PrimaryButton
        menuProps={menuProps}
        menuAs={getMenu}
        iconProps={{ iconName: "MultiSelect" }}
        text={strings.TakeBulkActionsEntireLibrary}
        title={strings.TakeBulkActionsEntireLibraryTooltip}
        disabled={!hasItems || executingAction}
      />
    </Stack>
  );
};