import { ISpeaker } from "./models/ISpeaker";

export interface ISharePointServices {
    GetAllSpeakers(): Promise<ISpeaker[]>;
    GetAllSpeakersByCountry(country: string): Promise<ISpeaker[]>;
};