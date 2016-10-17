import { IUserConfig } from "./iUserConfig";
import { ISite } from "../types/iSite";

export interface IGlobalDeploymentConfig {
    User: IUserConfig;
    Sites: ISite[];
}
