export interface IRetentionLabel {
  name?: string;
  retentionSettings?: {
    behaviorDuringRetentionPeriod: string;
    isDeleteAllowed: boolean;
    isRecordLocked: boolean;
    isMetadataUpdateAllowed: boolean;
    isContentUpdateAllowed: boolean;
    isLabelUpdateAllowed: boolean;
  };
  isLabelAppliedExplicitly?: boolean;
  labelAppliedDateTime?: string;
  labelAppliedBy?: {
    user: {
      id: string;
      displayName: string;
    };
  };
}
