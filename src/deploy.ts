/// <reference path="../typings/index.d.ts" />

import * as FileSystem from "fs";
import * as minimist from "minimist";
import * as promptly from "promptly";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { MyConsoleLogger } from "./Logger/MyConsoleLogger";
import { GlobalDeploymentConfig } from "./Interfaces/Config/GlobalDeploymentConfig";
import { SiteDeploymentConfig } from "./Interfaces/Config/SiteDeploymentConfig";
import { DeploymentManager } from "./DeploymentManager";

function processGlobalDeploymentConfig(globalDeploymentConfig: GlobalDeploymentConfig) {
    if (globalDeploymentConfig.Sites && globalDeploymentConfig.Sites instanceof Array && globalDeploymentConfig.Sites.length > 0) {
        globalDeploymentConfig.Sites.forEach((siteCollection, index, array) => {
            //todo: create child process
            let siteDeploymentConfig: SiteDeploymentConfig = {
                User: globalDeploymentConfig.User,
                Site: siteCollection,
            };
            let deploymentManager = new DeploymentManager(siteDeploymentConfig);
            deploymentManager.deploy();
        });
    } else {
        Logger.write("None sites defined in deployment config.", Logger.LogLevel.Info);
    }
}

interface ConsoleArguments {
    deploymentConfigPath: string;
    userPassword: string;
    logLevel: string;
}

let clArgs: ConsoleArguments = <any>minimist(global.process.argv.slice(2), {
    alias: {
        f: "deploymentConfigPath",
        l: "logLevel",
    },
    default: {
        l: "Verbose",
    },
    string: ["f", "l"],
});

Logger.subscribe(new MyConsoleLogger());
Logger.activeLogLevel = Logger.LogLevel[clArgs.logLevel];

Logger.write("Start deployment", Logger.LogLevel.Info);
Logger.write(`Console arguments: ${JSON.stringify(clArgs)}`, 0);

if (clArgs.deploymentConfigPath) {
    let globalDeploymentConfig: GlobalDeploymentConfig = JSON.parse(FileSystem.readFileSync(clArgs.deploymentConfigPath, "utf8"));
    if (globalDeploymentConfig) {
        Logger.write(`Loaded deployment config: ${clArgs.deploymentConfigPath} `);
        Logger.write(JSON.stringify(globalDeploymentConfig), 0);

        Logger.write(`Authentication details:`);
        Logger.write(`Authtype: ${globalDeploymentConfig.User.authtype}.`);
        Logger.write(`Username: ${globalDeploymentConfig.User.username}.`);

        if (!globalDeploymentConfig.User.password) {
            promptly.password("Password:", (error, password) => {
                if (password) {
                    globalDeploymentConfig.User.password = password;
                    processGlobalDeploymentConfig(globalDeploymentConfig);
                } else {
                    throw new Error(`Requesting user password failed. ${error}`);
                }
            });
        } else {
            processGlobalDeploymentConfig(globalDeploymentConfig);
        }
    }
} else {
    Logger.write("The required deploy config path paramater is not available!", Logger.LogLevel.Error);
    Logger.write("deploy.js parameters are:");
    Logger.write("\tf - for deployment config path. (required)");
    Logger.write("\tl - for custom log level. Available log levels are Verbose, Info, Warning, Error and Off.");
}
