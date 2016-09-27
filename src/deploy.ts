/// <reference path="../typings/index.d.ts" />

import * as FileSystem from "fs";
import * as minimist from "minimist";
import * as promptly from "promptly";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { MyConsoleLogger } from "./Logger/MyConsoleLogger";
import { DeploymentConfig } from "./Interfaces/Config/DeploymentConfig";
import { DeploymentManager } from "./DeploymentManager";

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
    let deploymentConfig: DeploymentConfig = JSON.parse(FileSystem.readFileSync(clArgs.deploymentConfigPath, "utf8"));
    if (deploymentConfig) {
        Logger.write(`Loaded deployment config: ${clArgs.deploymentConfigPath} `);
        Logger.write(JSON.stringify(deploymentConfig), 0);

        Logger.write(`Authentication details:`);
        Logger.write(`Authtype: ${deploymentConfig.User.authtype}.`);
        Logger.write(`Username: ${deploymentConfig.User.username}.`);

        if (!deploymentConfig.User.password) {
            promptly.password("Password:", (error, password) => {
                if (password) {
                    deploymentConfig.User.password = password;
                    let deploymentManager = new DeploymentManager(deploymentConfig);
                    deploymentManager.deploy();
                } else {
                    throw new Error(`Requesting user password failed. ${error}`);
                }
            });
        } else {
            let deploymentManager = new DeploymentManager(deploymentConfig);
            deploymentManager.deploy();
        }
    }
} else {
    Logger.write("The required deploy config path paramater is not available!", Logger.LogLevel.Error);
    Logger.write("deploy.js parameters are:");
    Logger.write("\tf - for deployment config path. (required)");
    Logger.write("\tl - for custom log level. Available log levels are Verbose, Info, Warning, Error and Off.");
}
