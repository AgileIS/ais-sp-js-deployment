import { DeploymentManager } from "./DeploymentManager";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { MyConsoleLogger } from "./Logger/MyConsoleLogger";
import { SiteDeploymentConfig } from "./Interfaces/Config/SiteDeploymentConfig";
import { ForkProcessArguments } from "./Interfaces/Config/ForkProcessArguments";

let forkConfig: ForkProcessArguments = JSON.parse(process.argv[2]);
let siteDeploymentConfig: SiteDeploymentConfig = forkConfig.siteDeploymentConfig;
let logLevel: number = forkConfig.logLevel;

Logger.subscribe(new MyConsoleLogger(siteDeploymentConfig.Site.Url));
Logger.activeLogLevel = logLevel ? logLevel : Logger.LogLevel.Verbose;

try {
    let deploymentManager = new DeploymentManager(siteDeploymentConfig);
    deploymentManager.deploy();
} catch (error) {
    console.error("\x1b[31m", error);
}

