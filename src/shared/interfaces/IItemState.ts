
export interface IItemState {
  listItemId: number;
  toggling: boolean;
  clearing: boolean;
  errorToggling?: string | undefined;
  errorClearing: boolean;
}
