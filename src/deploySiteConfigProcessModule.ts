import { DeploymentManager } from "./deploymentManager";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { ConsoleLogger } from "./logger/consoleLogger";
import { FileLogger } from "./logger/fileLogger";
import { ISiteDeploymentConfig } from "./interfaces/config/iSiteDeploymentConfig";
import { IForkProcessArguments } from "./interfaces/config/iForkProcessArguments";

let forkConfig: IForkProcessArguments = JSON.parse(process.argv[2]);
let siteDeploymentConfig: ISiteDeploymentConfig = forkConfig.siteDeploymentConfig;
let logLevel: number = forkConfig.logLevel;

Logger.subscribe(new ConsoleLogger(process.pid.toString()));
Logger.subscribe(new FileLogger(siteDeploymentConfig.Site.Url));
Logger.activeLogLevel = logLevel ? logLevel : Logger.LogLevel.Verbose;

try {
    let deploymentManager = new DeploymentManager(siteDeploymentConfig);
    deploymentManager.deploy();
} catch (error) {
    console.error("\x1b[31m", error);
}
