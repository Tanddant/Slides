import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISharePointServices } from "./ISharePointServices";
import { SPHttpClient } from "@microsoft/sp-http";
import { ISpeaker } from "./models/ISpeaker";

export class SharePointServices3 implements ISharePointServices {
    private context: WebPartContext;

    constructor(context: WebPartContext) {
        this.context = context;
    }

    public async GetAllSpeakers(): Promise<ISpeaker[]> {
        return this.SPGetListItemsRequest("/_api/web/lists/getbytitle('Speakers')/items");
    }

    public async GetAllSpeakersByCountry(country: string): Promise<ISpeaker[]> {
        return this.SPGetListItemsRequest("/_api/web/lists/getbytitle('Speakers')/items", `Country eq '${country}'`);
    }

    private async SPGetListItemsRequest(relativePath: string, filter?: string): Promise<any[]> {
        let url: URL = new URL(relativePath, this.context.pageContext.web.absoluteUrl);
        if (filter)
            url.searchParams.append("$filter", filter);


        const response = await this.context.spHttpClient.get(url.toString(), SPHttpClient.configurations.v1)
        if (response.ok) {
            const result = await response.json();
            return result.d.results;
        } else {
            throw new Error(`Failed to fetch data from SharePoint: ${response.statusText}`);
        }
    }

}