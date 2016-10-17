import { IUserConfig } from "./iUserConfig";
import { ISite } from "../types/iSite";

export interface ISiteDeploymentConfig {
    User: IUserConfig;
    Site: ISite;
}
