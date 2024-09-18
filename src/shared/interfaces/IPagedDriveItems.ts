import { IDriveItem } from "./IDriveItem";

export interface IPagedDriveItems {
    nextLink?: string;
    items: IDriveItem[];
}
  