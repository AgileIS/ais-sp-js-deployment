import { UserConfig } from "./UserConfig";
import { SiteCollectionConfig } from "./SiteCollectionConfig";

export interface DeploymentConfig {
    userConfig: UserConfig;
    siteCollectionConfigs: SiteCollectionConfig[];
}