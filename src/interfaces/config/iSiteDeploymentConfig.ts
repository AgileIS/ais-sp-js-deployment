import { IUserConfig } from "./iUserConfig";
import { ISiteConfig } from "./iSiteConfig";

export interface ISiteDeploymentConfig {
    User: IUserConfig;
    Site: ISiteConfig;
}
