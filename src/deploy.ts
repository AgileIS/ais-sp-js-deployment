/// <reference path="../typings/index.d.ts" />

import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { MyConsoleLogger } from "./Logger/MyConsoleLogger";
import { DeploymentConfig } from "./interface/Config/DeploymentConfig";
import * as FileSystem from "fs";
import * as minimist from "minimist";

interface ConsoleArguments {
    deploymentConfigPath: string;
    userPassword: string;
    logLevel: Logger.LogLevel;
}

let clArgs: ConsoleArguments = <any> minimist(global.process.argv.slice(2), {
    alias: {
        f: "deploymentConfigPath",
        l: "logLevel",
        p: "userPassword",
    },
    default: {
        l: Logger.LogLevel.Verbose
    },
    string: ["f", "p", "l"]
});

Logger.subscribe(new MyConsoleLogger());
Logger.activeLogLevel = clArgs.logLevel;

Logger.write("Start deployment", Logger.LogLevel.Info);
Logger.write(`Console arguments: ${JSON.stringify(clArgs)}`, 0);

if (clArgs.deploymentConfigPath && clArgs.userPassword) {
    let deploymentConfig: DeploymentConfig = JSON.parse(FileSystem.readFileSync(clArgs.deploymentConfigPath, "utf8"));
    if (deploymentConfig) {
        Logger.write(`Loaded deployment config: ${clArgs.deploymentConfigPath} `);
        Logger.write(JSON.stringify(deploymentConfig), 0);

    }
} else {
    Logger.write("Required deploy paramater are not available!", Logger.LogLevel.Error);
}
