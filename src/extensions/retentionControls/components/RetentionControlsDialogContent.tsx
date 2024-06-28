import { DefaultButton, DialogContent, DialogFooter, DialogType, format, IDialogFooterStyles, IMessageBarStyles, IStackItemStyles, IStackTokens, Link, mergeStyles, mergeStyleSets, MessageBar, MessageBarType, PrimaryButton, ResponsiveMode, Shimmer, Spinner, SpinnerSize, Stack } from "@fluentui/react";
import { BaseComponentContext } from "@microsoft/sp-component-base"
import * as React from "react";
import { SharePointService } from "../../../shared/services/SharePointService";
import { IRetentionLabel } from "../../../shared/interfaces/IRetentionLabel";
import { useState } from "react";
import * as strings from "RetentionControlsCommandSetStrings";
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { FontIcon } from '@fluentui/react/lib/Icon';
import { Warning } from "../../../shared/Warning";
initializeIcons();

export interface IRetentionControlsDialogProps {
    context: BaseComponentContext;
    listId: string;
    listItemIds: number[];
    close: { (): void }
}

const stackItemStyles: IStackItemStyles = {
    root: {
      alignItems: 'center',
      display: 'flex',
      width: '250px'
    },
};

const dialogFooterStyles: IDialogFooterStyles = {     
    action: {
      width: '100%'
    },
    actions: {},
    actionsRight: {}
};

const messageBarStyles: IMessageBarStyles = {     
    root: {
      marginBottom: '10px'
    }
};

const stackTokens: IStackTokens = {
    childrenGap: 5,    
};

const iconClass = mergeStyles({
    fontSize: 14,
    height: 14,
    width: 14,
    margin: '0 10px 0 0',
});

const classNames = mergeStyleSets({
    green: [{ color: 'darkgreen' }, iconClass],
    red: [{ color: 'indianred' }, iconClass],
    blue: [{ color: '#28a8ea' }, iconClass]
});

const getBehaviorLabel = (behavior: string | undefined): string => { 
    switch (behavior) {
        case "retain":
            return "Retain";
        case "doNotRetain":
            return "Do not retain";
        case "retainAsRecord":
            return "Retain as record";
        case "retainAsRegulatoryRecord":
            return "Retain as regulatory record";
        default:
            return "N/A";
    }
}

const RetentionControlsDialogContent: React.FC<IRetentionControlsDialogProps> = (props) => {
    const [error, setError] = useState<string | undefined>();
    const [warning, setWarning] = useState<string | undefined>();
    const [loading, setLoading] = useState<boolean>(true);
    const [successMessage, setSuccessMessage] = useState<string>();
    const [clearing, setClearing] = useState<boolean>(false);
    const [toggling, setToggling] = useState<boolean>(false);
    const [driveItemLabel, setDriveItemLabel] = useState<IRetentionLabel | undefined>();
    const spoService = props.context.serviceScope.consume(SharePointService.serviceKey);

    const clearLabel = async (): Promise<void> => {        
        setSuccessMessage(undefined);
        setError(undefined);
        setWarning(undefined);

        try {
            if (driveItemLabel?.retentionSettings?.isRecordLocked === false) {
                throw new Error(strings.CannotClearWhileUnlocked);
            }

            setClearing(true);
            await spoService.clearRetentionLabels(props.listId, props.listItemIds);
            setDriveItemLabel(undefined);
            setSuccessMessage(strings.LabelCleared);
            setClearing(false);
        }
        catch(error) {
            if ((error as Warning).isWarning) {
                setWarning(error.message);
            } else {
                setError(error.message);
            }
            setClearing(false);
        }
    }

    const toggleLockStatus = async (): Promise<void> => {
        setSuccessMessage(undefined);
        setError(undefined);
        setWarning(undefined);

        if (driveItemLabel?.driveId === undefined || driveItemLabel?.driveItemId === undefined) {
            return;
        }

        try {
            setToggling(true);
            const newDriveItemLabel = await spoService.toggleLockStatus(driveItemLabel.driveId, driveItemLabel.driveItemId, driveItemLabel?.retentionSettings?.isRecordLocked === true ? false : true);
            setDriveItemLabel(newDriveItemLabel);
            setSuccessMessage(format(strings.RecordStatusToggled, newDriveItemLabel?.retentionSettings?.isRecordLocked === true ? strings.Locked.toLowerCase() : strings.Unlocked.toLowerCase()));
            setToggling(false);
        }
        catch(error) {
            setError(error.message);
            setToggling(false);
        }
    }

    React.useEffect(() => {
        const fetchData = async (): Promise<void> => {
            try {
                const response = await spoService.getRetentionSettings(props.listId, props.listItemIds[0]);
                setDriveItemLabel(response);
                setLoading(false);
            }
            catch(error) {
                setError(error.message);
                setLoading(false);
            }
        }

        if (props.listItemIds.length === 1) {
            fetchData().catch(() => { setError(error) });
        }
        else {
            setLoading(false);
        }
    }, []);

    const applied = driveItemLabel?.labelAppliedDateTime ? new Date(driveItemLabel.labelAppliedDateTime).toLocaleDateString() : "N/A"

    return (
        <DialogContent styles={{ content: { maxWidth: '600px'}}} type={DialogType.largeHeader} responsiveMode={ResponsiveMode.small} showCloseButton={true} title={strings.RetentionControlsHeader} onDismiss={props.close}>
            {
                successMessage ?
                    <MessageBar styles={messageBarStyles} messageBarType={MessageBarType.success}>
                        {successMessage}
                    </MessageBar> :
                <></>
            }
            {
                error ? 
                    <MessageBar styles={messageBarStyles} messageBarType={MessageBarType.error}>
                        {error} 
                    </MessageBar> :
                <></>
            }
            {
                warning ? 
                    <MessageBar styles={messageBarStyles} messageBarType={MessageBarType.warning}>
                        {warning} 
                    </MessageBar> :
                <></>
            }
            {
                !loading && props.listItemIds.length === 1 && driveItemLabel?.name === undefined ?
                    <MessageBar styles={messageBarStyles} style={{ width: '500px'}}>
                        {strings.NoLabelApplied}
                    </MessageBar> :
                <></>
            }
            {
                !loading && props.listItemIds.length > 1 ?
                    <MessageBar styles={messageBarStyles} style={{ width: '500px'}}>
                        {format(strings.MultipleItemsSelected, props.listItemIds.length)}
                    </MessageBar> :
                <></>
            }
            {
                props.listItemIds.length === 1 && 
                    (loading || 
                    !loading && driveItemLabel?.name !== undefined) ?
                        <Stack tokens={{childrenGap: 10}}>
                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={100} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.RetentionLabelApplied}</strong>
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={250} height={16} isDataLoaded={!loading}>
                                        <div>
                                            {driveItemLabel?.name}
                                        </div>
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>          
                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={50} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.RetentionLabelApplicationDate}</strong>
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={250} height={16} isDataLoaded={!loading}>
                                        <div>
                                            {applied}
                                        </div>
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>  
                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={70} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.RetentionLabelAppliedBy}</strong>
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={250} height={16} isDataLoaded={!loading}>
                                        <div>
                                            {driveItemLabel?.labelAppliedBy?.user?.displayName}
                                        </div>
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>

                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={220} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.BehaviorDuringRetentionPeriod}</strong>
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={150} height={16} isDataLoaded={!loading}>
                                        {getBehaviorLabel(driveItemLabel?.retentionSettings?.behaviorDuringRetentionPeriod)}
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>
                            
                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={170} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.IsMetadataUpdateAllowed}</strong> 
                                        <FontIcon style={{ marginLeft: '6px', cursor: 'pointer'}} iconName="Info" title={strings.IsMetadataUpdateAllowedTooltip} className={classNames.blue} />
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={50} height={16} isDataLoaded={!loading}>
                                        { driveItemLabel?.retentionSettings?.isMetadataUpdateAllowed === true ? <><FontIcon iconName="Accept" className={classNames.green} /> {strings.ToggleOnText}</> : <><FontIcon className={classNames.red} iconName="Cancel" /> {strings.ToggleOffText}</> }
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>
                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={160} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.IsContentUpdateAllowed}</strong> 
                                        <FontIcon style={{ marginLeft: '6px', cursor: 'pointer'}} iconName="Info" title={strings.IsContentUpdateAllowedTooltip} className={classNames.blue} />
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={50} height={16} isDataLoaded={!loading}>
                                        { driveItemLabel?.retentionSettings?.isContentUpdateAllowed === true ? <><FontIcon iconName="Accept" className={classNames.green} /> {strings.ToggleOnText}</> : <><FontIcon className={classNames.red} iconName="Cancel" /> {strings.ToggleOffText}</> }
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>
                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={100} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.IsDeleteAllowed}</strong>
                                        <FontIcon style={{ marginLeft: '6px', cursor: 'pointer'}} iconName="Info" title={strings.IsDeleteAllowedTooltip} className={classNames.blue} />
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={50} height={16} isDataLoaded={!loading}>
                                        { driveItemLabel?.retentionSettings?.isDeleteAllowed === true ? <><FontIcon iconName="Accept" className={classNames.green} /> {strings.ToggleOnText}</> : <><FontIcon className={classNames.red} iconName="Cancel" /> {strings.ToggleOffText}</> }
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>
                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={140} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.IsLabelUpdateAllowed}</strong>
                                        <FontIcon style={{ marginLeft: '6px', cursor: 'pointer'}} iconName="Info" title={strings.isLabelUpdateAllowedTooltip} className={classNames.blue} />
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={50} height={16} isDataLoaded={!loading}>
                                        { driveItemLabel?.retentionSettings?.isLabelUpdateAllowed === true ? <><FontIcon iconName="Accept" className={classNames.green} /> {strings.ToggleOnText}</> : <><FontIcon className={classNames.red} iconName="Cancel" /> {strings.ToggleOffText}</> }
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>
                            
                            <Stack horizontal tokens={stackTokens}>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={100} height={16} isDataLoaded={!loading}>
                                        <strong>{strings.RecordStatus}</strong>
                                        <FontIcon style={{ marginLeft: '6px', cursor: 'pointer'}} iconName="Info" title={strings.RecordStatusTooltip} className={classNames.blue} />
                                    </Shimmer>
                                </Stack.Item>
                                <Stack.Item grow={1} styles={stackItemStyles}>
                                    <Shimmer width={50} height={16} isDataLoaded={!loading}>
                                        <>
                                            { driveItemLabel?.retentionSettings?.isRecordLocked === true ? 
                                                <><FontIcon iconName="LockSolid" /> {strings.Locked}</> : 
                                                <><FontIcon iconName="Unlock" /> {strings.Unlocked}</> }
                                            <Link disabled={toggling} onClick={toggleLockStatus} style={{marginLeft: '10px'}}>
                                                { toggling ? <Spinner size={SpinnerSize.xSmall} style={{marginRight:'10px'}} labelPosition="right" label={strings.Toggling} /> : <>{strings.ToggleLockStatus}</> }
                                                
                                            </Link>  
                                        </>
                                    </Shimmer>
                                </Stack.Item>
                            </Stack>
                        </Stack> :
                    <></>
            }
                    
            <DialogFooter styles={dialogFooterStyles}>
                <Stack horizontal enableScopedSelectors horizontalAlign="space-between">
                    {
                        loading || (!loading && props.listItemIds.length > 0) ? 
                            <DefaultButton onClick={clearLabel} disabled={loading}>
                                { clearing ? <Spinner size={SpinnerSize.xSmall} style={{marginRight:'10px'}} /> : <></> }
                                {strings.ClearLabel}
                            </DefaultButton> :
                        <></>
                    }
                    
                    <PrimaryButton onClick={props.close} text={strings.CloseModal} />
                </Stack>
            </DialogFooter>
        </DialogContent>
    )
}
export default RetentionControlsDialogContent;