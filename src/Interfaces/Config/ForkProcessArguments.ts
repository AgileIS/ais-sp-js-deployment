import { ISiteDeploymentConfig } from "./SiteDeploymentConfig";

export interface IForkProcessArguments {
    siteDeploymentConfig: ISiteDeploymentConfig;
    logLevel: number;
}
