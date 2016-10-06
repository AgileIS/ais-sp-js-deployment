import { DeploymentManager } from "./DeploymentManager";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { MyConsoleLogger } from "./Logger/MyConsoleLogger";

let forkConfig = JSON.parse(process.argv[2]);
let siteDeploymentConfig = forkConfig.siteDeploymentConfig;
let logLevel = forkConfig.logLevel;

Logger.subscribe(new MyConsoleLogger());
//string to int
Logger.activeLogLevel = Logger.LogLevel.Verbose;

let deploymentManager = new DeploymentManager(siteDeploymentConfig);
deploymentManager.deploy();
