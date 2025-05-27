import { ISharePointServices } from "./ISharePointServices";
import { SPFI } from "@pnp/sp/presets/all";
import { ISpeaker } from "./models/ISpeaker";

export class SharePointServices4 implements ISharePointServices {
    private SPFI: SPFI;

    constructor(SPFi: SPFI) {
        this.SPFI = SPFi;
    }

    public async GetAllSpeakers(): Promise<ISpeaker[]> {
        return this.SPFI.web.lists.getByTitle("Speakers").items();
    }

    public async GetAllSpeakersByCountry(country: string): Promise<ISpeaker[]> {
        return this.SPFI.web.lists.getByTitle("Speakers").items.filter(`Country eq '${country}'`)();
    }
}

