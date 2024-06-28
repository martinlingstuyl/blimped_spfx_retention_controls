declare interface IRetentionControlsCommandSetStrings {
  RetentionControlsHeader: string;
  RetentionLabelApplied: string;
  RetentionLabelApplicationDate: string;
  RetentionLabelAppliedBy: string;
  RecordStatus: string;
  IsDeleteAllowed: string;
  BehaviorDuringRetentionPeriod: string;
  IsMetadataUpdateAllowed: string;
  IsContentUpdateAllowed: string;
  IsLabelUpdateAllowed: string;
  ToggleOnText: string;
  ToggleOffText: string;
  Locked: string;
  Unlocked: string;
  Toggling: string;
  ToggleLockStatus: string;
  BehaviorRetain: string;
  BehaviorDoNotRetain: string;
  BehaviorRetainAsRecord: string;
  BehaviorRetainAsRegulatoryRecord: string;
  NoLabelApplied: string;
  MultipleItemsSelected: string;
  ClearLabel: string;
  LabelCleared: string;
  RecordStatusToggled: string;
  CloseModal: string;
  CannotClearWhileUnlocked: string;
  ClearErrorForMultipleItems: string;
  ClearErrorForSingleItem: string;
  UnhandledError: string;
  IsMetadataUpdateAllowedTooltip: string;
  IsDeleteAllowedTooltip: string;
  isLabelUpdateAllowedTooltip: string;
  IsContentUpdateAllowedTooltip: string;
  RecordStatusTooltip: string;
}

declare module 'RetentionControlsCommandSetStrings' {
  const strings: IRetentionControlsCommandSetStrings;
  export = strings;
}
