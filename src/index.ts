/// <reference path="../typings/index.d.ts" />

import * as FileSystem from "fs";
import * as promptly from "promptly";
import * as childProcess from "child_process";
import * as path from "path";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { MyConsoleLogger } from "./Logger/MyConsoleLogger";
import { IGlobalDeploymentConfig } from "./Interfaces/Config/GlobalDeploymentConfig";
import { IForkProcessArguments } from "./Interfaces/Config/ForkProcessArguments";

export namespace AisDeploy {

    let processCount = 0;

    function onChildProcessExit(code: number, signal: string): void {
        Logger.write("child ends " + this.pid, Logger.LogLevel.Info);
        processCount--;
        if (processCount === 0) {
            Logger.write("all site collections processed", Logger.LogLevel.Info);
        }
    }

    function onChildProcessDisconnect(): void {
        Logger.write("child disconnect " + this.pid, Logger.LogLevel.Info);
    }

    function processGlobalDeploymentConfig(globalDeploymentConfig: IGlobalDeploymentConfig, loglevel: Logger.LogLevel, runChildProcessInhDebugMode: boolean) {
        if (globalDeploymentConfig.Sites && globalDeploymentConfig.Sites instanceof Array && globalDeploymentConfig.Sites.length > 0) {
            globalDeploymentConfig.Sites.forEach((siteCollection, index, array) => {
                let forkOptions: childProcess.ForkOptions = { silent: false };
                if (runChildProcessInhDebugMode) {
                    forkOptions.execArgv = [`--debug-brk=5858${index}`];
                }

                let forkArgs: IForkProcessArguments = {
                    siteDeploymentConfig: {
                        User: globalDeploymentConfig.User,
                        Site: siteCollection,
                    },
                    logLevel: loglevel,
                };
                let child = childProcess.fork(__dirname + path.sep + "DeploySiteConfigProcessModule.js", [JSON.stringify(forkArgs)], forkOptions);
                child.on("disconnect", onChildProcessDisconnect);
                child.on("exit", onChildProcessExit);
                processCount++;
            });
        } else {
            Logger.write("None sites defined in deployment config.", Logger.LogLevel.Info);
        }
    }

    export function deploy(deploymentConfigPath: string, logLevel?: string, runChildProcessInhDebugMode?: boolean): void {
        Logger.subscribe(new MyConsoleLogger());
        Logger.activeLogLevel = Logger.LogLevel[logLevel ? logLevel : "Verbose"];

        Logger.write("Start deployment", Logger.LogLevel.Info);

        if (deploymentConfigPath) {
            let globalDeploymentConfig: IGlobalDeploymentConfig = JSON.parse(FileSystem.readFileSync(deploymentConfigPath, "utf8"));
            if (globalDeploymentConfig) {
                Logger.write(`Loaded deployment config: ${deploymentConfigPath} `);
                // Logger.write(JSON.stringify(globalDeploymentConfig), 0);

                Logger.write(`Authentication details:`);
                Logger.write(`Authtype: ${globalDeploymentConfig.User.authtype}.`);
                Logger.write(`Username: ${globalDeploymentConfig.User.username}.`);

                if (!globalDeploymentConfig.User.password) {
                    promptly.password("Password:", (error, password) => {
                        if (password) {
                            globalDeploymentConfig.User.password = password;
                            processGlobalDeploymentConfig(globalDeploymentConfig, Logger.activeLogLevel, runChildProcessInhDebugMode);
                        } else {
                            throw new Error(`Requesting user password failed. ${error}`);
                        }
                    });
                } else {
                    processGlobalDeploymentConfig(globalDeploymentConfig, Logger.activeLogLevel, runChildProcessInhDebugMode);
                }
            }
        } else {
            Logger.write("The required deploy config path paramater is not available!", Logger.LogLevel.Error);
            Logger.write("deploy.js parameters are:");
            Logger.write("\tf - for deployment config path. (required)");
            Logger.write("\tl - for custom log level. Available log levels are Verbose, Info, Warning, Error and Off.");
        }
    }
}
