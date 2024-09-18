import { IDriveItem } from "./interfaces/IDriveItem";
import { IItemMetadata } from "./interfaces/IItemMetadata";

export const getBehaviorLabel = (behavior: string | undefined): string => {
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
};

export const isRecordTypeLabel = (behavior: string | undefined): boolean => {
  return behavior === "retainAsRecord" || behavior === "retainAsRegulatoryRecord";
}

export const flattenItemMetadata = (item: IDriveItem | undefined): IItemMetadata | undefined => {
  if (item === undefined) {
    return undefined;
  }

  const siteRelativeFolder = item.parentReference.path.split(":").pop()?.replace(/^\//,'');
  const libraryRelativePath = siteRelativeFolder ? `${siteRelativeFolder}/${item.name}` : item.name;
  
  return { 
    id: parseFloat(item.listItem.id),
    driveItemId: item.id,
    name: item.name,
    path: libraryRelativePath,
    contentTypeId: item.listItem.contentType.id,
    isFolder: item.listItem.contentType.id.indexOf("0x0120") !== -1,
    isRecordTypeLabel: isRecordTypeLabel(item.retentionLabel?.retentionSettings?.behaviorDuringRetentionPeriod),
    retentionLabel: item.retentionLabel?.name,
    labelAppliedBy: item?.retentionLabel?.labelAppliedBy?.user?.displayName || (item?.retentionLabel?.labelAppliedBy as { application?: { displayName: string } })?.application?.displayName,
    labelAppliedDate: item?.retentionLabel?.labelAppliedDateTime ? new Date(item?.retentionLabel?.labelAppliedDateTime).toLocaleDateString() : undefined,
    eventDate: item?.listItem.fields.TagEventDate !== undefined && item?.listItem.fields.TagEventDate?.indexOf("9999") === -1 ? new Date(item?.listItem.fields.TagEventDate).toLocaleDateString() : undefined,
    behaviorDuringRetentionPeriod: item.retentionLabel?.retentionSettings?.behaviorDuringRetentionPeriod,
    isDeleteAllowed: item.retentionLabel?.retentionSettings?.isDeleteAllowed,
    isRecordLocked: item.retentionLabel?.retentionSettings?.isRecordLocked,
    isMetadataUpdateAllowed: item.retentionLabel?.retentionSettings?.isMetadataUpdateAllowed,
    isContentUpdateAllowed: item.retentionLabel?.retentionSettings?.isContentUpdateAllowed,
    isLabelUpdateAllowed: item.retentionLabel?.retentionSettings?.isLabelUpdateAllowed
  } as IItemMetadata;
};

export const flattenItemMetadataList = (items: IDriveItem[] | undefined): IItemMetadata[] => {
  return items ? items.filter(i => i.retentionLabel?.name !== undefined).map(item => flattenItemMetadata(item) as IItemMetadata) : [];
};

export const updateObjectProperties = <T>(original: T, updated: T): void => {
  for (const key in updated) {
    if (Object.prototype.hasOwnProperty.call(original, key)) {
      (original as never)[key] = (updated as never)[key];
    }
  }
};