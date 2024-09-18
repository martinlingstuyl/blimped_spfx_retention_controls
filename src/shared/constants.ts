import * as strings from "RetentionControlsCommandSetStrings";
import { ICustomColumn } from "./interfaces/ICustomColumn";

export const itemMetadataColumns: ICustomColumn[] = [
  { key: "icon", name: "Icon", fieldName: "icon", minWidth: 16, maxWidth: 16, isResizable: false, isIconOnly: true, iconName: "Page"  },
  { key: "name", name: "Name", fieldName: "name", minWidth: 80, maxWidth: 200, isResizable: true, },
  { key: "retentionLabel", name: "Label", fieldName: "retentionLabel", minWidth: 80, maxWidth: 200, isResizable: true },
  { key: "labelAppliedBy", name: "Applied by", fieldName: "labelAppliedBy", minWidth: 80, maxWidth: 200, isResizable: true },
  { key: "labelAppliedDate", name: "Applied", fieldName: "labelAppliedDate", minWidth: 80, maxWidth: 200, isResizable: true },
  { key: "eventDate", name: "Event date", fieldName: "eventDate", minWidth: 80, maxWidth: 200, isResizable: true },
  { key: "behaviorDuringRetentionPeriod", name: "Behavior", fieldName: "name", minWidth: 80, maxWidth: 200, isResizable: true },
  { key: "isDeleteAllowed", name: "Delete", fieldName: "isDeleteAllowed", minWidth: 16, maxWidth: 16, isResizable: false, isIconOnly: true, iconName: "Delete", title: strings.IsDeleteAllowed },
  { key: "isMetadataUpdateAllowed", name: "Metadata update", fieldName: "isMetadataUpdateAllowed", minWidth: 16, maxWidth: 16, isResizable: false, isIconOnly: true, iconName: "PageHeaderEdit", title: strings.IsMetadataUpdateAllowed },
  { key: "isContentUpdateAllowed", name: "Content update", fieldName: "isContentUpdateAllowed", minWidth: 16, maxWidth: 16, isResizable: false, isIconOnly: true, iconName: "PageEdit", title: strings.IsContentUpdateAllowed },
  { key: "isLabelUpdateAllowed", name: "Label update", fieldName: "isLabelUpdateAllowed", minWidth: 16, maxWidth: 16, isResizable: false, isIconOnly: true, iconName: "Tag", title: strings.IsLabelUpdateAllowed },
  { key: "isRecordLocked", name: "Locked", fieldName: "isRecordLocked", minWidth: 16, maxWidth: 16, isResizable: false, isIconOnly: true, iconName: "Lock", title: strings.RecordStatus },
  { key: "clearLabel", name: "clearLabel", fieldName: "clearLabel", minWidth: 16, maxWidth: 16, isResizable: false, isIconOnly: true, iconName: "Untag", title: strings.ClearLabel },
];