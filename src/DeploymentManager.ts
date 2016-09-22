import * as PNP from "@agileis/sp-pnp-js";
import { LibraryConfiguration } from "@agileis/sp-pnp-js/lib/configuration/pnplibconfig";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { DeploymentConfig } from "./interface/Config/DeploymentConfig";
import { SiteCollectionConfig } from "./interface/Config/SiteCollectionConfig";
import { ISPObjectHandler } from "./interface/ObjectHandler/ispobjecthandler";
import { SiteHandler } from "./ObjectHandler/SiteHandler";
import { ListHandler } from "./ObjectHandler/ListHandler";
import { FieldHandler } from "./ObjectHandler/FieldHandler";
import { ViewHandler } from "./ObjectHandler/ViewHandler";
import { ViewFieldHandler } from "./ObjectHandler/ViewFieldHandler";
import { ContentTypeHandler } from "./ObjectHandler/ContentTypeHandler";
import { AuthenticationType } from "./Constants/AuthenticationType";
import { NodeHttpProxy } from "./NodeHttpProxy";
import * as url from "url";
import { NodeJsomHandler } from "./NodeJsomHandler";

export class DeploymentManager {
    private _deploymentConfig: DeploymentConfig;
    private _objectHandlers: { [id: string]: ISPObjectHandler } = {
        Sites: new SiteHandler(),
        ContentTypes: new ContentTypeHandler(),
        Lists: new ListHandler(),
        Fields: new FieldHandler(),
        Views: new ViewHandler(),
        ViewFields: new ViewFieldHandler(),
    };
    private _deployDependencies: Promise<void>;

    constructor(deploymentConfig: DeploymentConfig) {
        this._deploymentConfig = deploymentConfig;
        this.setupProxy();
        this.setupPnPJs();
        this._deployDependencies = NodeJsomHandler.initialize(deploymentConfig);
    }

    public deploy(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            this._deployDependencies
                .then( response => {
                    this.processConfig(this._deploymentConfig)
                        .then(() => {
                            Logger.write("All site collection processed", Logger.LogLevel.Info);
                            resolve();
                        })
                        .catch((error) => {
                            Logger.write("Error occured while processing site collections - " + error, Logger.LogLevel.Info);
                            reject(error);
                        });
                })
                .catch(error => {
                    reject(error);
                });
        });
    }

    private processConfig(config: any, parentPromise?: Promise<any>): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            let processingPromises: Array<Promise<any>> = [];

            Object.keys(config).forEach((nodeKey, nodeIndex) => {
                Logger.write("Processing node " + nodeKey + " at index " + nodeIndex, 0);
                let handler = this._objectHandlers[nodeKey];
                if (handler) {
                    Logger.write(`Found handler ${handler.constructor.name} for node ${nodeKey}`, Logger.LogLevel.Verbose);
                    if (config[nodeKey] instanceof Array) {
                        config[nodeKey].forEach(subNode => {
                            Logger.write("Call the handler " + handler.constructor.name + " for the node:" + JSON.stringify(subNode), Logger.LogLevel.Verbose);
                            let handlerPromise = handler.execute(subNode, parentPromise);
                            processingPromises.push(handlerPromise);
                            processingPromises = processingPromises.concat(this.processConfig(subNode, handlerPromise));
                        });
                    } else {
                        Logger.write("Call the handler " + handler.constructor.name + " for the node:" + JSON.stringify(config[nodeKey]), Logger.LogLevel.Verbose);
                        let handlerPromise = handler.execute(config[nodeKey], parentPromise);
                        processingPromises.push(handlerPromise);
                        processingPromises = processingPromises.concat(this.processConfig(config[nodeKey], handlerPromise));
                    }
                }
            });

            Promise.all(processingPromises)
                .then(() => {
                    resolve();
                })
                .catch((error) => {
                    reject(error);
                });
        });
    }

    private setupProxy(): void {
        if (this._deploymentConfig.User.proxyUrl) {
            NodeHttpProxy.url = url.parse(this._deploymentConfig.User.proxyUrl);
            NodeHttpProxy.activate();
        }
    }

    private setupPnPJs(): void {
        let userConfig = this._deploymentConfig.User;
        Logger.write("Setup pnp-core-js", Logger.LogLevel.Info);
        Logger.write(`pno-core-js authentication type: ${userConfig.authtype}`, Logger.LogLevel.Info);

        let pnpConfig: LibraryConfiguration;
        if (String(userConfig.authtype).toLowerCase() === AuthenticationType.Ntlm.toLowerCase()) {
            let userAndDommain = userConfig.username.split("\\");
            if (!userConfig.workstation) {
                userConfig.workstation = "";
            }

            pnpConfig = {
                nodeHttpNtlmClientOptions: {
                    domain: userAndDommain[0],
                    password: userConfig.password,
                    siteUrl: "",
                    username: userAndDommain[1],
                    workstation: userConfig.workstation,
                },
            };
        } else if (String(userConfig.authtype).toLowerCase() === AuthenticationType.Basic.toLowerCase()) {
            pnpConfig = {
                nodeHttpBasicClientOptions: {
                    password: userConfig.password,
                    siteUrl: "",
                    username: userConfig.username,
                },
            };
        } else {
            throw new Error(`Unsupported authentication type. Use ${AuthenticationType.Ntlm} or ${AuthenticationType.Basic} `);
        }

        if (pnpConfig) {
            PNP.setup(pnpConfig);
        }
    }
}
