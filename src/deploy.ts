/// <reference path="../typings/index.d.ts" />

import * as FileSystem from "fs";
import * as minimist from "minimist";
import * as promptly from "promptly";
import * as childProcess from "child_process";
import * as path from "path";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { MyConsoleLogger } from "./Logger/MyConsoleLogger";
import { GlobalDeploymentConfig } from "./Interfaces/Config/GlobalDeploymentConfig";
import { ForkProcessArguments } from "./Interfaces/Config/ForkProcessArguments";

let processCount = 0;
function onChildProcessExit(code: number, signal: string): void {
    Logger.write("child ends " + this.pid, Logger.LogLevel.Info);
    processCount--;
}

function onChildProcessDisconnect(): void {
    Logger.write("child disconnect " + this.pid + " killing...", Logger.LogLevel.Info);
}

function processGlobalDeploymentConfig(globalDeploymentConfig: GlobalDeploymentConfig, loglevel: Logger.LogLevel) {
    if (globalDeploymentConfig.Sites && globalDeploymentConfig.Sites instanceof Array && globalDeploymentConfig.Sites.length > 0) {
        globalDeploymentConfig.Sites.forEach((siteCollection, index, array) => {
            // let forkOptions: childProcess.ForkOptions = { silent: false, execArgv: ["--debug-brk=58589"] }; //for debugging
            let forkOptions: childProcess.ForkOptions = { silent: false };
            let forkArgs: ForkProcessArguments = {
                siteDeploymentConfig: {
                    User: globalDeploymentConfig.User,
                    Site: siteCollection,
                },
                logLevel: loglevel,
            };
            // todo: kill child process on end and kill all child process on parent child process kill
            let child = childProcess.fork(__dirname + path.sep + "DeploySiteConfigProcessModule.js", [JSON.stringify(forkArgs)], forkOptions);
            child.on("disconnect", onChildProcessDisconnect);
            child.on("exit", onChildProcessExit);
            processCount++;
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
        // Logger.write(JSON.stringify(globalDeploymentConfig), 0);

        Logger.write(`Authentication details:`);
        Logger.write(`Authtype: ${globalDeploymentConfig.User.authtype}.`);
        Logger.write(`Username: ${globalDeploymentConfig.User.username}.`);

        if (!globalDeploymentConfig.User.password) {
            promptly.password("Password:", (error, password) => {
                if (password) {
                    globalDeploymentConfig.User.password = password;
                    processGlobalDeploymentConfig(globalDeploymentConfig, Logger.activeLogLevel);
                } else {
                    throw new Error(`Requesting user password failed. ${error}`);
                }
            });
        } else {
            processGlobalDeploymentConfig(globalDeploymentConfig, Logger.activeLogLevel);
        }
    }
} else {
    Logger.write("The required deploy config path paramater is not available!", Logger.LogLevel.Error);
    Logger.write("deploy.js parameters are:");
    Logger.write("\tf - for deployment config path. (required)");
    Logger.write("\tl - for custom log level. Available log levels are Verbose, Info, Warning, Error and Off.");
}
