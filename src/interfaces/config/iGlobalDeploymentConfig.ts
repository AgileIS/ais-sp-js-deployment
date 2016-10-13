import { IUserConfig } from "./iUserConfig";
import { ISiteConfig } from "./iSiteConfig";

export interface IGlobalDeploymentConfig {
    User: IUserConfig;
    Sites: ISiteConfig[];
}
