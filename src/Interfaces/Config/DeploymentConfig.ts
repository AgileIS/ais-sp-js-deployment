import { UserConfig } from "./UserConfig";
import { SiteCollectionConfig } from "./SiteCollectionConfig";

export interface DeploymentConfig {
    User: UserConfig;
    Sites: SiteCollectionConfig[];
}
