import { SiteDeploymentConfig } from "./SiteDeploymentConfig";

export interface ForkProcessArguments {
    siteDeploymentConfig: SiteDeploymentConfig;
    logLevel: number;
}
