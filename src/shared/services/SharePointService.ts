import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { IRetentionLabel } from "../interfaces/IRetentionLabel";
import { Warning } from "../Warning";
import * as strings from "RetentionControlsCommandSetStrings";
import { format } from "@fluentui/react";
import { IListItemFields } from "../interfaces/IListItemFields";

export interface ISharePointService {
  getListItemFields: (listId: string, listItemId: number) => Promise<IListItemFields>;
  getRetentionSettings: (listId: string, listItemId: number) => Promise<IRetentionLabel | undefined>;
  clearRetentionLabels: (listItemId: number[]) => Promise<void>;
  toggleLockStatus: (listItemId: string, lockStatus: boolean) => Promise<void>;
}

export class SharePointService implements ISharePointService {
  public static readonly serviceKey: ServiceKey<ISharePointService> = ServiceKey.create<ISharePointService>("SPFx:SharePointService", SharePointService);

  private _pageContext: PageContext;
  private _spoHttpClient: SPHttpClient;

  public constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(async () => {
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._spoHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
    });
  }

  public async getListItemFields(listId: string, listItemId: number): Promise<IListItemFields> {
    const requestUrl = `${this._pageContext.site.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})?$select=TagEventDate`;
    const response = await this._spoHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();
      throw new Error(error?.error?.message ?? strings.UnhandledError);
    }

    const responseContent: IListItemFields = await response.json();
    return responseContent;
  }

  public async getRetentionSettings(listId: string, listItemId: number): Promise<IRetentionLabel> {
    const driveId = await this.getDriveId(listId);
    const driveItemId = await this.getDriveItemId(driveId, listItemId);

    const siteUrl = new URL(this._pageContext.site.absoluteUrl);
    const requestUrl = `${siteUrl.protocol}//${siteUrl.host}/_api/v2.0/drives/${driveId}/items/${driveItemId}/retentionLabel`;

    const response = await this._spoHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();
      throw new Error(error?.error?.message ?? strings.UnhandledError);
    }

    const responseContent: IRetentionLabel = await response.json();
    return responseContent;
  }

  public async clearRetentionLabels(listItemIds: number[]): Promise<void> {
    if (this._pageContext.list === undefined) {
      throw new Error("List information not available");
    }
    const listAbsoluteUrl = new URL(this._pageContext.list.serverRelativeUrl, this._pageContext.site.absoluteUrl);

    const url = `${this._pageContext.site.absoluteUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`;

    const body = {
      listUrl: listAbsoluteUrl.href,
      complianceTagValue: "",
      itemIds: listItemIds,
    };

    const response = await this._spoHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();
      throw new Error(error?.error?.message ?? strings.UnhandledError);
    }

    const content: { value?: number[] } = await response.json();

    if (content.value && content.value.length > 0) {
      if (listItemIds.length !== content.value.length) {
        throw new Warning(format(strings.ClearErrorForMultipleItems, content.value.length, listItemIds.length));
      }

      throw new Error(strings.ClearErrorForSingleItem);
    }
  }

  public async toggleLockStatus(listItemId: string, lockStatus: boolean): Promise<void> {
    if (this._pageContext.list === undefined) {
      throw new Error("List information not available");
    }

    const url = lockStatus ? `${this._pageContext.site.absoluteUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.LockRecordItem()` : `${this._pageContext.site.absoluteUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`;

    const body = {
      listUrl: this._pageContext.list.serverRelativeUrl,
      itemId: listItemId,
    };

    const response = await this._spoHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();
      throw new Error(error?.error?.message ?? strings.UnhandledError);
    }
  }

  private async getDriveId(listId: string): Promise<string> {
    const url = new URL(this._pageContext.site.absoluteUrl);
    const cacheKey = `Blimped_RC_driveId_${this._pageContext.site.id}_${this._pageContext.web.id}_${listId}`;
    const cachedDriveId = sessionStorage.getItem(cacheKey);

    if (cachedDriveId !== null) {
      return cachedDriveId;
    }

    const siteId = `${url.hostname},${this._pageContext.site.id},${this._pageContext.web.id}`;
    const siteUrl = new URL(this._pageContext.site.absoluteUrl);
    const requestUrl = `${siteUrl.protocol}//${siteUrl.host}/_api/v2.0/sites/${siteId}/lists/${listId}?$expand=drive($select=id)&$select=id`;
    const response = await this._spoHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();
      throw new Error(error?.error?.message ?? strings.UnhandledError);
    }

    const responseContent: { drive: { id: string } } = await response.json();
    sessionStorage.setItem(cacheKey, responseContent.drive.id);

    return responseContent.drive.id;
  }

  private async getDriveItemId(driveId: string, listItemId: number): Promise<string> {
    const siteUrl = new URL(this._pageContext.site.absoluteUrl);
    const requestUrl = `${siteUrl.protocol}//${siteUrl.host}/_api/v2.0/drives/${driveId}/items?$filter=sharepointIds/listItemId eq '${listItemId}'&$select=sharepointIds,id`;
    const response = await this._spoHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();
      throw new Error(error?.error?.message ?? strings.UnhandledError);
    }

    const responseContent: { value: { id: string }[] } = await response.json();

    return responseContent.value[0].id;
  }
}
