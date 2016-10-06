import { UserConfig } from "./UserConfig";
import { SiteConfig } from "./SiteConfig";

export interface SiteDeploymentConfig {
    User: UserConfig;
    Site: SiteConfig;
}
