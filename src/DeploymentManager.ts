import * as PnP from "@agileis/sp-pnp-js";
import { LibraryConfiguration } from "@agileis/sp-pnp-js/lib/configuration/pnplibconfig";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { DeploymentConfig } from "./Interfaces/Config/DeploymentConfig";
import { ISPObjectHandler } from "./Interfaces/ObjectHandler/ISPObjectHandler";
import { ISPObjectHandlerCollection } from "./Interfaces/ObjectHandler/ISPObjectHandlerCollection";
import { IList } from "./Interfaces/Types/IList";
import { IPromiseResult } from "./Interfaces/IPromiseResult";
import { SiteHandler } from "./ObjectHandler/SiteHandler";
import { ListHandler } from "./ObjectHandler/ListHandler";
import { ItemHandler } from "./ObjectHandler/ItemHandler";
import { FileHandler } from "./ObjectHandler/FileHandler";
import { FieldHandler } from "./ObjectHandler/FieldHandler";
import { ViewHandler } from "./ObjectHandler/ViewHandler";
import { SolutionHandler } from "./ObjectHandler/SolutionHandler";
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
    private _deployDependencies: Array<Promise<any>> = new Array();
    private _objectHandlers: ISPObjectHandlerCollection = {
        Features: new FeatureHandler(),
        Sites: new SiteHandler(),
        ContentTypes: new ContentTypeHandler(),
        Lists: new ListHandler(),
        Fields: new FieldHandler(),
        Views: new ViewHandler(),
        Items: new ItemHandler(),
        Navigation: new NavigationHandler(),
        Files: new FileHandler(),
        Solutions: new SolutionHandler(),
    };

    constructor(deploymentConfig: DeploymentConfig) {
        if (deploymentConfig.Sites && deploymentConfig.Sites.length === 1) {
            this._deploymentConfig = <DeploymentConfig>JSON.parse(
                Util.replaceUrlTokens(JSON.stringify(deploymentConfig), Util.getRelativeUrl(deploymentConfig.Sites[0].Url), `_layouts/${deploymentConfig.Sites[0].LayoutsHive}`));
            this.setupProxy();
            this.setupPnPJs();
            this._deployDependencies.push(NodeJsomHandler.initialize(deploymentConfig));
        } else {
            throw new Error("Deployment config site count is not equals 1");
        }
    }

    public deploy(): Promise<void> {
        return Promise.all(this._deployDependencies).then(() => {
            return this.processDeploymentConfig()
                .then(() => {
                    Logger.write("All site collection processed", Logger.LogLevel.Info);
                })
                .catch((error) => {
                    Logger.write("Error occured while processing site collections - " + Util.getErrorMessage(error), Logger.LogLevel.Info);
                });
        });
    }

    private processDeploymentConfig(): Promise<any> {
        let siteProcessingPromise = this._objectHandlers.Sites.execute(this._deploymentConfig.Sites[0], Promise.resolve());

        let nodeProcessingOrder: string[] = ["Features", "Fields", "ContentTypes", "Lists", "Navigation", "Files", "Solutions"];
        let existingSiteNodes = Object.keys(this._deploymentConfig.Sites[0]);

        return nodeProcessingOrder.reduce((dependentPromise, processingKey, proecssingIndex, array): Promise<any> => {
            return dependentPromise
                .then(() => {
                    let processingConfig = (<any>this._deploymentConfig.Sites[0])[processingKey];
                    let processingHandler = this._objectHandlers[processingKey];
                    let prossingPromise: Promise<any> = Promise.resolve();

                    if (existingSiteNodes.indexOf(processingKey) > -1 && processingHandler) {
                        if (processingKey === "Fields" || processingKey === "Files") {
                            prossingPromise = this.processDeploymentConfigNodesParallel(processingHandler, processingConfig, siteProcessingPromise);
                        } else if (processingKey === "Features" || processingKey === "ContentTypes" || processingKey === "Solutions") {
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

    private processListsDeploymentConfig(listProcessingHandler: ISPObjectHandler, listsDeploymentConfig: IList[], sitePromise: Promise<IPromiseResult<Web>>): Promise<any> {
        let listPromiseDictionary: { [internalName: string]: Promise<any> } = {};
        let listProcessingPromises: Promise<any>[] = new Array();

        listsDeploymentConfig.forEach((listConfig, index, array) => {
            let processingPromise = listProcessingHandler.execute(listConfig, sitePromise);
            listPromiseDictionary[listConfig.InternalName] = processingPromise;
            listProcessingPromises.push(processingPromise);
        });

        return Promise.all(listProcessingPromises).then(() => {
            return listsDeploymentConfig.reduce((dependentPromise, listConfig, listIndex, listsArray) => {
                let listPromise = listPromiseDictionary[listConfig.InternalName];
                return dependentPromise
                    .then(() => {
                        return this.processDeploymentConfigNodesSequential(this._objectHandlers.Fields, listConfig.Fields, listPromise);
                    })
                    .then(() => {
                        return Promise.all([
                            this.processDeploymentConfigNodesParallel(this._objectHandlers.Views, listConfig.Views, listPromise),
                            this.processDeploymentConfigNodesParallel(this._objectHandlers.Items, listConfig.Items, listPromise),
                            this.processDeploymentConfigNodesParallel(this._objectHandlers.Files, listConfig.Files, listPromise)]
                        );
                    });
            }, Promise.resolve());
        });
    };

    private processDeploymentConfigNodesParallel(processingHandler: ISPObjectHandler, deploymentConfigNodes: Array<any>, dependentPromise: Promise<any>): Promise<any> {
        let processingPromisses: Array<Promise<any>> = [Promise.resolve()];

        if (processingHandler && deploymentConfigNodes && deploymentConfigNodes instanceof Array && deploymentConfigNodes.length > 0) {
            deploymentConfigNodes.forEach(
                (processingConfig, proecssingIndex, array) => {
                    processingPromisses.push(processingHandler.execute(processingConfig, dependentPromise));
                });
        } if (!processingHandler) {
            Logger.write("Processing object handler is undefined while processing deployment config nodes parallel.", Logger.LogLevel.Error);
            processingPromisses.push(Promise.reject(undefined));
        }

        return Promise.all(processingPromisses);
    }

    private processDeploymentConfigNodesSequential(processingHandler: ISPObjectHandler, deploymentConfigNodes: Array<any>, dependentPromise: Promise<any>): Promise<any> {
        let processingPromise: Promise<any> = Promise.resolve();

        if (processingHandler && deploymentConfigNodes && deploymentConfigNodes instanceof Array && deploymentConfigNodes.length > 0) {
            processingPromise = deploymentConfigNodes.reduce(
                (previousPromise, processingConfig, proecssingIndex, array) => {
                    return previousPromise.then(() => {
                        return processingHandler.execute(processingConfig, dependentPromise);
                    });
                }, dependentPromise);
        } if (!processingHandler) {
            Logger.write("Processing object handler is undefined while processing deployment config nodes sequential.", Logger.LogLevel.Error);
            processingPromise = Promise.reject(undefined);
        }

        return processingPromise;
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
