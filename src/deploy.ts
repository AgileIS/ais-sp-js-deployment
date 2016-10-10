import * as minimist from "minimist";
import { AisDeploy } from "./index";

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

if (clArgs.deploymentConfigPath) {
    AisDeploy.deploy(clArgs.deploymentConfigPath, clArgs.logLevel);
} else {
    console.error("ERROR: missing arguments (deploymentConfigPath)!");
}
