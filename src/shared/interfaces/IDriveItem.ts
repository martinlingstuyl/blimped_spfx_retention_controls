import { IListItemFields } from "./IListItemFields";
import { IRetentionLabel } from "./IRetentionLabel";

export interface IDriveItem {
  name: string;
  id: string;
  parentReference: { path: string };
  listItem: { fields: IListItemFields, id: string, contentType: { id: string, name: string } };
  retentionLabel?: IRetentionLabel;  
}