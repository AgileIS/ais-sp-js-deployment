import * as minimist from "minimist";
import { AisDeploy } from "./index";

interface IConsoleArguments {
    deploymentConfigPath: string;
    userPassword: string;
    logLevel: string;
    runChildProcessInhDebugMode: boolean;
}

let clArgs: IConsoleArguments = <any>minimist(global.process.argv.slice(2), {
    alias: {
        f: "deploymentConfigPath",
        l: "logLevel",
        d: "runChildProcessInhDebugMode",
    },
    default: {
        l: "Verbose",
        d: false,
    },
    string: ["f", "l"],
    boolean: ["d"],
});

if (clArgs.deploymentConfigPath) {
    AisDeploy.deploy(clArgs.deploymentConfigPath, clArgs.logLevel, clArgs.runChildProcessInhDebugMode);
} else {
    console.error("\x1b[31m", "ERROR: missing arguments (deploymentConfigPath)!");
}
