import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientV3, MSGraphClientFactory, SPHttpClient } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { IRetentionLabel } from "../interfaces/IRetentionLabel";
import { Warning } from "../Warning";
import * as strings from "RetentionControlsCommandSetStrings";
import { format } from "@fluentui/react";

export interface ISharePointService {
    getRetentionSettings: (listId: string, listItemId: number) => Promise<IRetentionLabel>;
    clearRetentionLabels: (listId: string, listItemId: number[]) => Promise<void>;
    toggleLockStatus: (driveId: string, driveItemId: string, lockStatus: boolean) => Promise<IRetentionLabel>;
}

export class SharePointService implements ISharePointService {

    public static readonly serviceKey: ServiceKey<ISharePointService> =
        ServiceKey.create<ISharePointService>('SPFx:SharePointService', SharePointService);

    private _pageContext: PageContext;
    private _graphClientFactory: MSGraphClientFactory;
    private _spoHttpClient: SPHttpClient;

    public constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(async () => {
            this._graphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
            this._pageContext = serviceScope.consume(PageContext.serviceKey);        
            this._spoHttpClient = serviceScope.consume(SPHttpClient.serviceKey);        
        });
    }

    public async getRetentionSettings(listId: string, listItemId: number): Promise<IRetentionLabel> {
        const graphClient = await this._graphClientFactory.getClient("3");        
        const driveId = await this.getDriveId(graphClient, listId);
        const driveItemId = await this.getDriveItemId(graphClient, driveId, listItemId);

        const response: IRetentionLabel = await graphClient.api(`/drives/${driveId}/items/${driveItemId}/retentionLabel`).get();

        if (!response) {
            throw new Error(strings.UnhandledError);
        }

        response.driveId = driveId;
        response.driveItemId = driveItemId;

        return response;
    }

    public async clearRetentionLabels(listId: string, listItemIds: number[]): Promise<void> {
        if (this._pageContext.list === undefined) {
            throw new Error('List information not available');
        }
        const listAbsoluteUrl = new URL(this._pageContext.list.serverRelativeUrl, this._pageContext.site.absoluteUrl);

        const url = `${this._pageContext.site.absoluteUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`;

        const body = {
            listUrl: listAbsoluteUrl.href,
            complianceTagValue: "",
            itemIds: listItemIds,
        }

        const response = await this._spoHttpClient.post(url, SPHttpClient.configurations.v1, { 
            body: JSON.stringify(body),
        });
        
        if (!response) {
            throw new Error(strings.UnhandledError);
        }

        const content: { value?: number[] } = await response.json();

        if (content.value && content.value.length > 0) {
            if (listItemIds.length !== content.value.length) {
                throw new Warning(format(strings.ClearErrorForMultipleItems, content.value.length, listItemIds.length));
            }

            throw new Error(strings.ClearErrorForSingleItem);
        }
    }

    public async toggleLockStatus(driveId: string, driveItemId: string, lockStatus: boolean): Promise<IRetentionLabel> {
        if (this._pageContext.list === undefined) {
            throw new Error('List information not available');
        }
        const graphClient = await this._graphClientFactory.getClient("3");        

        const response: IRetentionLabel = await graphClient.api(`/drives/${driveId}/items/${driveItemId}/retentionLabel`).patch({
            "retentionSettings": {
              "isRecordLocked": lockStatus
            }
        });

        response.driveId = driveId;
        response.driveItemId = driveItemId;

        return response;
    }

    private async getDriveId(graphClient: MSGraphClientV3, listId: string): Promise<string> {
        const url = new URL(this._pageContext.site.absoluteUrl);
        
        const cachedDriveId = sessionStorage.getItem(`SPFX_RetentionControls_driveId_${this._pageContext.site.absoluteUrl}_${listId}`);

        if (cachedDriveId !== null) {
            return cachedDriveId;
        }
        
        const siteId = `${url.hostname},${this._pageContext.site.id},${this._pageContext.web.id}`
        
        const response: { drive: { id: string } } | undefined = await graphClient.api(`/sites/${siteId}/lists/${listId}?$expand=drive($select=id)&$select=id`).get();
        
        if (!response) {
            throw new Error(strings.UnhandledError);
        }

        sessionStorage.setItem(`SPFX_RetentionControls_driveId_${this._pageContext.site.absoluteUrl}_${listId}`, response.drive.id)

        return response.drive.id;
    }
    
    private async getDriveItemId(graphClient: MSGraphClientV3, driveId: string, listItemId: number): Promise<string> {
        const response:{ value: { id: string }[] } | undefined = await graphClient.api(`/drives/${driveId}/items?$filter=sharepointIds/listItemId eq '${listItemId}'&$select=sharepointIds,id`).get();

        if (!response || !response.value || response.value.length === 0) {
            throw new Error(strings.UnhandledError);
        }

        return response.value[0].id;
    }
}