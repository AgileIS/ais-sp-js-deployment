import { IUserConfig } from "./UserConfig";
import { ISiteConfig } from "./SiteConfig";

export interface ISiteDeploymentConfig {
    User: IUserConfig;
    Site: ISiteConfig;
}
