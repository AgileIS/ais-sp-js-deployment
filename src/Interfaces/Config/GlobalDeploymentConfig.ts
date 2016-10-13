import { IUserConfig } from "./UserConfig";
import { ISiteConfig } from "./SiteConfig";

export interface IGlobalDeploymentConfig {
    User: IUserConfig;
    Sites: ISiteConfig[];
}
