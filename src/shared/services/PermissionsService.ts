import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory, SPHttpClient } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { IPermissionEntry, IUserPermissions } from "../interfaces/IPermissions";

export interface IPermissionsService {
  getUserPermissions(permissions: IPermissionEntry[]): Promise<IUserPermissions>;
}

export class PermissionsService implements IPermissionsService {
  public static readonly serviceKey: ServiceKey<IPermissionsService> = ServiceKey.create<IPermissionsService>("SPFx:PermissionsService", PermissionsService);

  private _pageContext: PageContext;
  private _msGraphClientFactory: MSGraphClientFactory;
  private _spHttpClient: SPHttpClient;

  public constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(async () => {
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
    });
  }

  private getSessionStorageKey(): string {
    return `retentionControls_${this._pageContext.site.absoluteUrl}`;
  }

  public async getUserPermissions(permissions: IPermissionEntry[]): Promise<IUserPermissions> {
    const storageKey = this.getSessionStorageKey();

    // Check if sessionStorage has cached permissions for this site
    try {
      const cachedPermissions = sessionStorage.getItem(storageKey);
      if (cachedPermissions) {
        const parsedPermissions = JSON.parse(cachedPermissions) as IUserPermissions;
        return parsedPermissions;
      }
    } catch (error) {
      console.warn("Retention Controls: Error reading cached permissions:", error);
    }

    // First, check SharePoint edit permissions on the site
    const hasSharePointEditPermissions = await this.checkSharePointEditPermissions();
    
    let userPermissions: IUserPermissions;
    
    // If user doesn't have SharePoint edit permissions, return false regardless of custom permissions
    if (!hasSharePointEditPermissions) {
      userPermissions = {
        editMode: false,
        hasSharePointEditPermissions: false
      };
    } else {
      // If user has SharePoint permissions, check custom permissions configuration
      let hasCustomEditPermissions = true; // Default to true if no custom permissions configured
      
      if (permissions && permissions.length > 0) {
        hasCustomEditPermissions = await this.validateUserPermissions(permissions);
      }
      
      userPermissions = {
        editMode: hasCustomEditPermissions,
        hasSharePointEditPermissions: true
      };
    }

    // Cache the permissions
    try {
      sessionStorage.setItem(storageKey, JSON.stringify(userPermissions));
    } catch (error) {
      console.warn("Retention Controls: Error caching permissions:", error);
    }

    return userPermissions;
  }

  private async checkSharePointEditPermissions(): Promise<boolean> {
    try {
      // Check user's effective permissions on the site
      // Using the current user's login name to check permissions
      const requestUrl = `${this._pageContext.web.absoluteUrl}/_api/web/getusereffectivepermissions(@v)?@v='${encodeURIComponent(`i:0#.f|membership|${this._pageContext.user.loginName}`)}'`;
      
      const response = await this._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
      
      if (!response.ok) {
        console.warn("Retention Controls: Failed to check SharePoint permissions:", response.statusText);
        return false;
      }

      const data = await response.json();
      
      // Check for EditListItems permission (0x2 = EditListItems in SharePoint)
      // SharePoint permissions are stored as a High/Low pair representing a 64-bit number
      // EditListItems is bit 1 (0x2), so we check if the Low part has bit 1 set
      const hasEditPermission = (data.Low & 2) === 2;
      
      return hasEditPermission;
    } catch (error) {
      console.error("Retention Controls: Error checking SharePoint edit permissions:", error);
      return false;
    }
  }

  private async validateUserPermissions(permissions: IPermissionEntry[]): Promise<boolean> {
    if (!permissions || permissions.length === 0) {
      return false;
    }

    const currentUserLoginName = this._pageContext.user.loginName;
    const currentUserEmail = this._pageContext.user.email;

    for (const permission of permissions) {
      // Check user UPN/email match
      if (permission.userName && (permission.userName.toLowerCase() === currentUserLoginName.toLowerCase() || 
                                 permission.userName.toLowerCase() === currentUserEmail.toLowerCase())) {
        return true;
      }

      // Check group membership
      if (permission.groupId) {
        const isGroupMember = await this.checkGroupMembership(permission.groupId);
        if (isGroupMember) {
          return true;
        }
      }
    }

    return false;
  }

  private async checkGroupMembership(groupId: string): Promise<boolean> {
    try {
      const msGraphClient = await this._msGraphClientFactory.getClient("3");

      // Check if current user is a member of the group
      const membershipResponse = await msGraphClient
        .api(`/me/transitiveMemberOf`)
        .filter(`id eq '${groupId}'`)
        .select("id")
        .get();

      return membershipResponse.value && membershipResponse.value.length > 0;
    } catch (error) {
      console.error("Retention Controls: Error checking group membership:", error);
      return false;
    }
  }
}
