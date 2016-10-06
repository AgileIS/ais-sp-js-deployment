import { UserConfig } from "./UserConfig";
import { SiteConfig } from "./SiteConfig";

export interface GlobalDeploymentConfig {
    User: UserConfig;
    Sites: SiteConfig[];
}
