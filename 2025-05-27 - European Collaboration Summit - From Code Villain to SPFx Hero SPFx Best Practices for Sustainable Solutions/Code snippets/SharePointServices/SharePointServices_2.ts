import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISharePointServices } from "./ISharePointServices";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export class SharePointServices2 implements ISharePointServices {
    private context: WebPartContext;

    constructor(context: WebPartContext) {
        this.context = context;
    }

    public async GetAllSpeakers(): Promise<any[]> {
        return this.SPGetRequest("/_api/web/lists/getbytitle('Speakers')/items");
    }

    public async GetAllSpeakersByCountry(country: string): Promise<any[]> {
        return this.SPGetRequest("/_api/web/lists/getbytitle('Speakers')/items?$filter=Country eq '" + country + "'");
    }

    private async SPGetRequest(url: string): Promise<any> {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + url, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            })
            .then((data: any) => {
                return data.value;
            });
    }
}