import { ISiteDeploymentConfig } from "./iSiteDeploymentConfig";

export interface IForkProcessArguments {
    siteDeploymentConfig: ISiteDeploymentConfig;
    logLevel: number;
}
