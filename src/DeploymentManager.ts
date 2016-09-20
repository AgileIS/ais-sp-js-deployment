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
import { NTLM } from "./ntlm";
import { SPJSOM } from "./node-spjsom";

class DeploymentManager {
    private _deploymentConfig: DeploymentConfig;
    private _objectHandlers: { [id: string]: ISPObjectHandler } = {
        Site: new SiteHandler(),
        ContentTypes: new ContentTypeHandler(),
        List: new ListHandler(),
        Field: new FieldHandler(),
        View: new ViewHandler(),
        ViewFields: new ViewFieldHandler(),
    };

    constructor(deploymentConfig: DeploymentConfig) {
        this._deploymentConfig = deploymentConfig;
    }

    public deploy(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            let processingPromises: Array<Promise<any>> = [];

            this._deploymentConfig.siteCollectionConfigs.forEach((siteCollectionConfig) => {
                processingPromises.push(this.processSiteCollection(siteCollectionConfig));
            });

            Promise.all(processingPromises)
                .then(() => {
                    Logger.write("All site collection processed", Logger.LogLevel.Info);
                    resolve();
                })
                .catch((error) => {
                    Logger.write("Error occured while processing site collections - " + error, Logger.LogLevel.Info);
                    reject(error);
                });
        });
    }

    private processSiteCollection(siteCollectionConfig: SiteCollectionConfig): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            let processingPromises: Array<Promise<any>> = [];
            
            Object.keys(siteCollectionConfig).forEach((nodeKey, nodeIndex) => {
                Logger.write("Processing node " + nodeKey + " at index " + nodeIndex, 0);
                let handler = this._objectHandlers[nodeKey];
                if (handler) {
                    Logger.write(`Found handler ${handler.constructor.name} for node ${nodeKey}`, Logger.LogLevel.Verbose);
                    if (siteCollectionConfig[nodeKey] instanceof Array) {
                        siteCollectionConfig[nodeKey].forEach(element => {
                            Logger.write("call object handler " + handler.constructor.name + " with element:" + JSON.stringify(element), 0);
                            let promise = handler.execute(element, parent);
                            processingPromises.push(promise);
                            processingPromises.concat(chooseAndUseHandler(element, promise));
                        });
                    } else {
                        Logger.write("call object handler " + handler.constructor.name + " with element:" + JSON.stringify(config[nodeKey]), 0);
                        let promise = handler.execute(config[nodeKey], parent);
                        processingPromises.push(promise);
                        processingPromises.concat(chooseAndUseHandler(config[nodeKey], promise));
                    }
                } else {
                    Logger.write(`Handler for node ${nodeKey} not available`, Logger.LogLevel.Warning);
                }
            });
        });
    }

    private setupProxy(): void {
        if (this._deploymentConfig.userConfig.proxyUrl) {
            NodeHttpProxy.url = url.parse("http://127.0.0.1:8888");
            NodeHttpProxy.activate();
        }
    }

    private setupSpJsom(): void {

    }

    private setupPnPJs(): void {
        let userConfig = this._deploymentConfig.userConfig;
        Logger.write("Setup pnp-core-js", Logger.LogLevel.Info);
        Logger.write(`pno-core-js authentication type: ${userConfig.authtype}`, Logger.LogLevel.Info);

        let pnpConfig: LibraryConfiguration;
        if (String(userConfig.authtype).toLowerCase() === AuthenticationType.Ntlm.toLowerCase()) {
            let userAndDommain = userConfig.username.split("\\");
            pnpConfig = {
                nodeHttpNtlmClientOptions: {
                    domain: userAndDommain[0],
                    password: userConfig.password,
                    siteUrl: "",
                    username: userConfig.username,
                    workstation: userConfig.workstation,
                },
            };
        } else if (String(userConfig.authtype).toLowerCase() === AuthenticationType.Basic.toLowerCase()) {
            pnpConfig = {
                nodeHttpBasicClientOptions: {
                    password: userConfig.password,
                    username: userConfig.username,
                    siteUrl: "",
                },
            };
        } else {
            throw new Error(`Unsupported authentication type. Use ${AuthenticationType.Ntlm} or ${AuthenticationType.Basic} `)
        }

        if (pnpConfig) {
            PNP.setup(pnpConfig);
        }
    }


}
