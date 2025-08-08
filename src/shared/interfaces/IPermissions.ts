export interface IPermissionEntry {
  /**
   * Entra ID Group ID (if checking group membership)
   */
  groupId?: string;
  
  /**
   * User UPN (if checking specific user)
   */
  userName?: string;
}

export interface IPermissions {
  /**
   * Array of permission entries that define who has edit access
   */
  entries: IPermissionEntry[];
}

export interface IUserPermissions {
  /**
   * Whether the user has edit permissions (can clear labels and toggle records)
   */
  editMode: boolean;
  
  /**
   * Whether the user has SharePoint edit permissions on the site
   */
  hasSharePointEditPermissions: boolean;
}
