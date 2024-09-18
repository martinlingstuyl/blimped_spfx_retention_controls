import { Guid, ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { IRetentionLabel } from "../interfaces/IRetentionLabel";
import * as strings from "RetentionControlsCommandSetStrings";
import { IListItemFields } from "../interfaces/IListItemFields";
import { IDriveItem as IDriveItem } from "../interfaces/IDriveItem";
import { IBatchItemResponse } from "../interfaces/IBatchErrorResponse";
import { IPagedDriveItems } from "../interfaces/IPagedDriveItems";

export interface ISharePointService {
  getPagedDriveItems(listId: string, pageSize?: number): Promise<IPagedDriveItems>
  getPagedDriveItemsUsingNextLink(nextLink: string): Promise<IPagedDriveItems>
  getDriveItems: (listId: string, listItemId: number[]) => Promise<IDriveItem[]>;
  getListItemFields: (listId: string, listItemId: number) => Promise<IListItemFields | undefined>;
  getRetentionSettings: (listId: string, listItemId: number) => Promise<IRetentionLabel | undefined>;
  clearRetentionLabels: (listItemIds: number[]) => Promise<IBatchItemResponse[]>;
  toggleLockStatus: (listItemIds: number[], lockStatus: boolean) => Promise<IBatchItemResponse[]>;
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

  /**
   * Get a recursively paged list of Drive Items with retention labels
   */
  public async getPagedDriveItems(listId: string, pageSize: number = 10): Promise<IPagedDriveItems> {
    const driveId = await this.getDriveId(listId);
    const requestUrl = `${this._pageContext.site.absoluteUrl}/_api/v2.0/drives/${driveId}/items`;
    
    const queryStrings = [
      `$filter=retentionLabel/name ne null`,
      `$expand=retentionLabel,listItem($select=id,contentType;$expand=fields($select=FileLeafRef,TagEventDate))`,
      `$select=id,name,parentReference,retentionLabel,id,listItem`,
      `$top=${pageSize}`
    ];

    return await this.executeGetPagedDriveItems(`${requestUrl}?${queryStrings.join("&")}`);    
  }

  public async getPagedDriveItemsUsingNextLink(nextLink: string): Promise<IPagedDriveItems> {    
    return await this.executeGetPagedDriveItems(nextLink);
  }

  /**
   * Get Drive Items based on a list of item ID's
   */
  public async getDriveItems(listId: string, listItemId: number[]): Promise<IDriveItem[]> {
    const driveId = await this.getDriveId(listId);
    const requestUrl = `${this._pageContext.site.absoluteUrl}/_api/v2.0/drives/${driveId}/items`;
    const filterString = listItemId.map((id) => `listItem/id eq '${id}'`).join(" or ");
    const queryStrings = [
      `$filter=${filterString}`,
      `$expand=retentionLabel,listItem($select=id,contentType;$expand=fields($select=FileLeafRef,TagEventDate))`,
      `$select=id,name,parentReference,retentionLabel,id,listItem`
    ];

    const response = await this._spoHttpClient.get(`${requestUrl}?${queryStrings.join("&")}`, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();
      throw new Error(error?.error?.message ?? strings.UnhandledError);
    }
    
    const responseContent: { value: IDriveItem[] } = await response.json();
    return responseContent.value;
  }

  public async getListItemFields(listId: string, listItemId: number): Promise<IListItemFields | undefined> {
    const requestUrl = `${this._pageContext.site.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${listItemId})?$select=TagEventDate`;
    const response = await this._spoHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();

      // Explainer: If the error "The field or property 'TagEventDate' does not exist." is returned,
      // it means the column is not present in the list because no event-based retention label
      // has been used. Just return nothing in that case.
      if (error?.error?.message.indexOf("TagEventDate") > -1) {
        return;
      }

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

  public async clearRetentionLabels(listItemIds: number[]): Promise<IBatchItemResponse[]> {
    if (this._pageContext.list === undefined) {
      throw new Error("List information not available");
    }

    const listAbsoluteUrl = new URL(this._pageContext.list.serverRelativeUrl, this._pageContext.site.absoluteUrl);

    const url = `${this._pageContext.site.absoluteUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`;

    const batchArray = [...listItemIds];    
    const allResponses: IBatchItemResponse[] = [];

    // Loop through the array in batches of 100 items, which is the max amount to post at this endpoint
    while (batchArray.length > 0) {
      const listItemIdsBatch = batchArray.splice(0, 100);
        
      const body = {
        listUrl: listAbsoluteUrl.href,
        complianceTagValue: "",
        itemIds: listItemIdsBatch,
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
        for (const itemId of listItemIdsBatch) {
          allResponses.push({ listItemId: itemId, success: content.value.indexOf(itemId) === -1 });
        }
      }
      else {
        for (const itemId of listItemIdsBatch) {
          allResponses.push({ listItemId: itemId, success: true });
        }
      }
    }

    return allResponses;    
  }

  public async toggleLockStatus(listItemIds: number[], lockStatus: boolean): Promise<IBatchItemResponse[]> {        
    const absoluteUrl = new URL(this._pageContext.site.absoluteUrl);
    const host = absoluteUrl.host;
    const requestUrl = `${this._pageContext.site.absoluteUrl}/_api/$batch`;

    const batchArray = [...listItemIds];
    const allResponses: IBatchItemResponse[] = [];

    // Loop through the array in batches of 100 items
    while (batchArray.length > 0) {
      const listItemIdsBatch = batchArray.splice(0, 100);
      const batchId = Guid.newGuid().toString();
      const body = this.buildBatchBody(listItemIdsBatch, lockStatus, batchId)

      const response = await this._spoHttpClient.post(requestUrl, SPHttpClient.configurations.v1, {
        body: body,
        headers: {
          "Content-Type": `multipart/mixed; boundary="batch_${batchId}"`,
          "Content-Transfer-Encoding": "binary",
          "Host": host
        },
      });

      if (!response.ok) {
        const error: { error: { message: string } } = await response.json();
        throw new Error(error?.error?.message ?? strings.UnhandledError);
      }

      const responseContent = await response.text();
      const responses = this.parseBatchResponseBody(responseContent, listItemIdsBatch);
      allResponses.push(...responses);
    }

    return allResponses;
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

  private buildBatchBody(listItemIds: number[], lockStatus: boolean, batchId: string): string {
    const serverRelativeUrl = this._pageContext.list?.serverRelativeUrl;
    if (serverRelativeUrl === undefined) {
      throw new Error("List information not available");
    }

    const batchUrl = lockStatus ? `${this._pageContext.site.absoluteUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.LockRecordItem()` : `${this._pageContext.site.absoluteUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.UnlockRecordItem()`;
    const changeSetId = Guid.newGuid().toString();
    const batchBody: string[] = [];

    batchBody.push(`--batch_${batchId}\n`);
    batchBody.push(`Content-Type: multipart/mixed; boundary="changeset_${changeSetId}"\n\n`);
    batchBody.push('Content-Transfer-Encoding: binary\n\n');
  
    listItemIds.forEach((listItemId, index) => {
      batchBody.push(`--changeset_${changeSetId}\n`);
      batchBody.push(`Content-Type: application/http\n`);
      batchBody.push(`Content-ID: ${index}\n`);
      batchBody.push(`Content-Transfer-Encoding: binary\n\n`);        
      batchBody.push(`POST ${batchUrl} HTTP/1.1\n`);
      batchBody.push(`Accept: application/json\n`);
      batchBody.push(`Content-Type: application/json;odata=nometadata\n\n`);        
      batchBody.push(`{ "listUrl": "${serverRelativeUrl}", "itemId": ${listItemId} }\n\n`);
      batchBody.push(``);        
    });

    batchBody.push(`--changeset_${changeSetId}--\n\n`);
    batchBody.push(`--batch_${batchId}--\n`);
    
    return batchBody.join('');
  }

  private parseBatchResponseBody(response: string, listItemIds: number[]): IBatchItemResponse[] {
    const responses: IBatchItemResponse[] = [];

    response.split('\r\n')
      .filter((line: string) => line.indexOf('{') === 0)
      .forEach((line: string, index: number) => {
        const parsedResponse = JSON.parse(line);

        if (parsedResponse.error) {
          // if an error object is returned, the request failed
          const error = parsedResponse.error as { message: string };
          responses.push({ errorMessage: error.message, listItemId: listItemIds[index], success: false });
        }
        else {
          responses.push({ listItemId: listItemIds[index], success: true });
        }
      });

    return responses;
  }
  
  private async executeGetPagedDriveItems(requestUrl: string): Promise<IPagedDriveItems> {
    const response = await this._spoHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      const error: { error: { message: string } } = await response.json();
      throw new Error(error?.error?.message ?? strings.UnhandledError);
    }
    
    const responseContent: { value: IDriveItem[] } = await response.json();
    return { nextLink: (responseContent as never)['@odata.nextLink'], items: responseContent.value };
  }
}
