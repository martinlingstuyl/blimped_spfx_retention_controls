export interface IItemMetadata {
  id: number;
  driveItemId: string;
  name: string;
  path: string;
  contentTypeId: string;
  isFolder: boolean;
  isRecordTypeLabel: boolean;
  retentionLabel?: string;
  labelAppliedBy?: string;
  labelAppliedDate?: string;
  eventDate?: string;
  behaviorDuringRetentionPeriod?: string;
  isDeleteAllowed?: boolean;
  isRecordLocked?: boolean;
  isMetadataUpdateAllowed?: boolean;
  isContentUpdateAllowed?: boolean;
  isLabelUpdateAllowed?: boolean;
}