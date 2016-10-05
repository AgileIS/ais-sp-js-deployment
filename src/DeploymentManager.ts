import * as PnP from "@agileis/sp-pnp-js";
import { LibraryConfiguration } from "@agileis/sp-pnp-js/lib/configuration/pnplibconfig";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { DeploymentConfig } from "./Interfaces/Config/DeploymentConfig";
import { ISPObjectHandler } from "./Interfaces/ObjectHandler/ISPObjectHandler";
import { IList } from "./Interfaces/Types/IList";
import { IPromiseResult } from "./Interfaces/IPromiseResult";
import { SiteHandler } from "./ObjectHandler/SiteHandler";
import { ListHandler } from "./ObjectHandler/ListHandler";
import { ItemHandler } from "./ObjectHandler/ItemHandler";
import { FileHandler } from "./ObjectHandler/FileHandler";
import { FieldHandler } from "./ObjectHandler/FieldHandler";
import { ViewHandler } from "./ObjectHandler/ViewHandler";
import { FeatureHandler } from "./ObjectHandler/FeatureHandler";
import { ContentTypeHandler } from "./ObjectHandler/ContentTypeHandler";
import { NavigationHandler } from "./ObjectHandler/NavigationHandler";
import { AuthenticationType } from "./Constants/AuthenticationType";
import { NodeHttpProxy } from "./NodeHttpProxy";
import { NodeJsomHandler } from "./NodeJsomHandler";
import { Util } from "./Util/Util";
import * as url from "url";

export class DeploymentManager {
    private _deploymentConfig: DeploymentConfig;
    private _objectHandlers: { [id: string]: ISPObjectHandler } = {
        Features: new FeatureHandler(),
        Sites: new SiteHandler(),
        ContentTypes: new ContentTypeHandler(),
        Lists: new ListHandler(),
        Fields: new FieldHandler(),
        Views: new ViewHandler(),
        Items: new ItemHandler(),
        Navigation: new NavigationHandler(),
        Files: new FileHandler(),
    };
    private _deployDependencies: Promise<void>;

    constructor(deploymentConfig: DeploymentConfig) {
        if (deploymentConfig.Sites && deploymentConfig.Sites.length === 1) {
            this._deploymentConfig = <DeploymentConfig>JSON.parse(
                Util.replaceUrlTokens(JSON.stringify(deploymentConfig), Util.getRelativeUrl(deploymentConfig.Sites[0].Url), `_layouts/${deploymentConfig.Sites[0].LayoutsHive}`));
            this.setupProxy();
            this.setupPnPJs();
            this._deployDependencies = NodeJsomHandler.initialize(deploymentConfig);
        } else {
            throw new Error("Deployment config site count is not equals 1");
        }
    }

    public deploy(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            this._deployDependencies
                .then(() => {
                    this.processDeploymentConfig()
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

    private processListsDeploymentConfig(listProcessingHandler: ISPObjectHandler, listsDeploymentConfig: IList[], sitePromise: Promise<IPromiseResult<Web>>): Promise<any> {
        let listPromiseDictionary: { [internalName: string]: Promise<any> } = {};
        let listProcessingPromises: Promise<any>[] = new Array();

        listsDeploymentConfig.forEach(
            (listConfig, index, array) => {
                let processingPromise = listProcessingHandler.execute(listConfig, sitePromise);
                listPromiseDictionary[listConfig.InternalName] = processingPromise;
                listProcessingPromises.push(processingPromise);
            }
        );

        return Promise.all(listsDeploymentConfig)
            .then(() => {
                return listsDeploymentConfig.reduce((dependentPromise, listConfig, listIndex, listsArray) => {
                    return dependentPromise.then(() => {
                        let fieldsProcessingPromise = undefined;
                        if (listConfig.Fields && listConfig.Fields instanceof Array && listConfig.Fields.length > 0) {
                            let fieldObjectHandlerKey = "Fields";
                            let fieldObjectHandler = this._objectHandlers[fieldObjectHandlerKey];
                            let listPromise = listPromiseDictionary[listConfig.InternalName];

                            fieldsProcessingPromise = listConfig.Fields.reduce((previousPromise, fieldConfig, fieldIndex, fieldsArray) => {
                                return previousPromise.then(() => {
                                    return fieldObjectHandler.execute(fieldConfig, listPromise);
                                });
                            }, listPromise);
                        } else {
                            fieldsProcessingPromise = Promise.resolve();
                        }

                        return fieldsProcessingPromise.then(() => {
                            //Views
                            return Promise.resolve();
                        });
                    });
                }, Promise.resolve());
            });
    };

    private processDeploymentConfigNodesParallel(processingHandler: ISPObjectHandler, deploymentConfigNodes: Array<any>, dependentPromise: Promise<any>): Promise<any> {
        let processingPromisses: Array<Promise<any>> = new Array();
        deploymentConfigNodes.forEach(
            (processingConfig, proecssingIndex, array) => {
                processingPromisses.push(processingHandler.execute(processingConfig, dependentPromise));
            });

        return Promise.all(processingPromisses);
    }

    private processDeploymentConfigNodesSequential(processingHandler: ISPObjectHandler, deploymentConfigNodes: Array<any>, dependentPromise: Promise<any>): Promise<any> {
        return deploymentConfigNodes.reduce(
            (previousPromise, processingConfig, proecssingIndex, array) => {
                return previousPromise.then(() => {
                    return processingHandler.execute(processingConfig, dependentPromise);
                });
            }, dependentPromise);
    }

    private processDeploymentConfig(): Promise<any> {
        let siteObjectHandlerKey = "Sites";
        let siteProcessingPromise: Promise<IPromiseResult<Web>> = this._objectHandlers[siteObjectHandlerKey].execute(this._deploymentConfig.Sites[0], Promise.resolve());

        let nodeProcessingOrder: string[] = ["Features", "Fields", "ContentTypes", "Lists", "Navigation", "Files"];
        let existingSiteNodes = Object.keys(this._deploymentConfig.Sites[0]);

        return nodeProcessingOrder.reduce(
            (dependentPromise, processingKey, proecssingIndex, array): Promise<any> => {
                return dependentPromise
                    .then(() => {
                        let processingConfig = (<any>this._deploymentConfig.Sites[0])[processingKey];
                        let processingHandler = this._objectHandlers[processingKey]
                        let prossingPromise: Promise<any> = undefined;

                        if ((existingSiteNodes.indexOf(processingKey) === -1)
                            || (processingConfig instanceof Array && processingConfig.length === 0)
                            || processingHandler === undefined) {
                            prossingPromise = Promise.resolve();
                        } else {
                            if (processingKey === "Fields" || processingKey === "Files") {
                                prossingPromise = this.processDeploymentConfigNodesParallel(processingHandler, processingConfig, siteProcessingPromise);
                            } else if (processingKey === "Features" || processingKey === "ContentTypes") {
                                prossingPromise = this.processDeploymentConfigNodesSequential(processingHandler, processingConfig, siteProcessingPromise);
                            } else if (processingKey === "Lists") {
                                prossingPromise = this.processListsDeploymentConfig(processingHandler, this._deploymentConfig.Sites[0].Lists, siteProcessingPromise);
                            } else if (processingKey === "Navigation") {
                                prossingPromise = processingHandler.execute(processingConfig, siteProcessingPromise);
                            }
                        }

                        return prossingPromise;
                    });
            }, siteProcessingPromise);
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
        Logger.write(`pnp-core-js authentication type: ${userConfig.authtype}`, Logger.LogLevel.Info);

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
                    siteUrl: this._deploymentConfig.Sites[0].Url,
                    username: userAndDommain[1],
                    workstation: userConfig.workstation,
                },
            };
        } else if (String(userConfig.authtype).toLowerCase() === AuthenticationType.Basic.toLowerCase()) {
            pnpConfig = {
                nodeHttpBasicClientOptions: {
                    password: userConfig.password,
                    siteUrl: this._deploymentConfig.Sites[0].Url,
                    username: userConfig.username,
                },
            };
        } else {
            throw new Error(`Unsupported authentication type. Use ${AuthenticationType.Ntlm} or ${AuthenticationType.Basic} `);
        }

        if (pnpConfig) {
            PnP.setup(pnpConfig);
        }
    }
}
