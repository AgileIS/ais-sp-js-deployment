import * as PnP from "@agileis/sp-pnp-js";
import { LibraryConfiguration } from "@agileis/sp-pnp-js/lib/configuration/pnplibconfig";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Folder } from "@agileis/sp-pnp-js/lib/sharepoint/rest/folders";
import { List } from "@agileis/sp-pnp-js/lib/sharepoint/rest/lists";
import { SiteDeploymentConfig } from "./Interfaces/Config/SiteDeploymentConfig";
import { ISPObjectHandler } from "./Interfaces/ObjectHandler/ISPObjectHandler";
import { ISPObjectHandlerCollection } from "./Interfaces/ObjectHandler/ISPObjectHandlerCollection";
import { IList } from "./Interfaces/Types/IList";
import { IFile } from "./Interfaces/Types/IFile";
import { IFolder } from "./Interfaces/Types/IFolder";
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
    private _siteDeploymentConfig: SiteDeploymentConfig;
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
    };

    constructor(siteDeploymentConfig: SiteDeploymentConfig) {
        if (siteDeploymentConfig.Site && siteDeploymentConfig.Site.Url) {
            this._siteDeploymentConfig = <SiteDeploymentConfig>JSON.parse(
                Util.replaceUrlTokens(JSON.stringify(siteDeploymentConfig), Util.getRelativeUrl(siteDeploymentConfig.Site.Url), `_layouts/${siteDeploymentConfig.Site.LayoutsHive}`));
            this.setupProxy();
            this.setupPnPJs();
            this._deployDependencies.push(NodeJsomHandler.initialize(siteDeploymentConfig));
        } else {
            throw new Error("Deployment config site or site url is undefined");
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
        let siteProcessingPromise = this._objectHandlers.Sites.execute(this._siteDeploymentConfig.Site, Promise.resolve());

        let nodeProcessingOrder: string[] = ["Features", "Fields", "ContentTypes", "Lists", "Navigation", "Files"];
        let existingSiteNodes = Object.keys(this._siteDeploymentConfig.Site);

        return nodeProcessingOrder.reduce((dependentPromise, processingKey, proecssingIndex, array): Promise<any> => {
            return dependentPromise
                .then(() => {
                    let processingConfig = (<any>this._siteDeploymentConfig.Site)[processingKey];
                    let processingHandler = this._objectHandlers[processingKey];
                    let processingPromise: Promise<any> = Promise.resolve();

                    if (existingSiteNodes.indexOf(processingKey) > -1 && processingHandler) {
                        if (processingKey === "Fields") {
                            processingPromise = this.processDeploymentConfigNodesParallel(processingHandler, processingConfig, siteProcessingPromise);
                        } else if (processingKey === "Features" || processingKey === "ContentTypes") {
                            processingPromise = this.processDeploymentConfigNodesSequential(processingHandler, processingConfig, siteProcessingPromise);
                        } else if (processingKey === "Lists") {
                            processingPromise = this.processListsDeploymentConfig(processingHandler, processingConfig, siteProcessingPromise);
                        } else if (processingKey === "Navigation") {
                            processingPromise = processingHandler.execute(processingConfig, siteProcessingPromise);
                        } else if (processingKey === "Files") {
                            processingPromise = this.processFilesDeploymentConfig(processingHandler, processingConfig, siteProcessingPromise);
                        }
                    }

                    return processingPromise;
                });
        }, siteProcessingPromise);
    }

    private processListsDeploymentConfig(listHandler: ISPObjectHandler, listsDeploymentConfig: IList[], sitePromise: Promise<IPromiseResult<Web>>): Promise<any> {
        let listPromiseDictionary: { [internalName: string]: Promise<any> } = {};
        let listProcessingPromises: Promise<any>[] = new Array();

        listsDeploymentConfig.forEach((listConfig, index, array) => {
            let processingPromise = listHandler.execute(listConfig, sitePromise);
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
                            this.processFilesDeploymentConfig(this._objectHandlers.Files, listConfig.Files, listPromise),
                        ]);
                    });
            }, Promise.resolve());
        });
    };

    private processFilesDeploymentConfig(filesHandler: ISPObjectHandler, filesDeploymentConfig: (IFile | IFolder)[], dependentPromise: Promise<IPromiseResult<Web | Folder | List>>) {
        let processingPromisses: Array<Promise<any>> = [Promise.resolve()];

        if (filesHandler && filesDeploymentConfig && filesDeploymentConfig instanceof Array && filesDeploymentConfig.length > 0) {
            filesDeploymentConfig.forEach((fileConfig, fileIndex, array) => {
                let fileProcessingPromise = filesHandler.execute(fileConfig, dependentPromise);
                if (Object.keys(fileConfig).indexOf("Files") > -1) {
                    let subFileProcessingPromise = fileProcessingPromise.then(() => {
                        return this.processFilesDeploymentConfig(filesHandler, (<IFolder>fileConfig).Files, fileProcessingPromise);
                    });
                    processingPromisses.push(subFileProcessingPromise);
                }
                processingPromisses.push(fileProcessingPromise);
            });
        } if (!filesHandler) {
            Logger.write("Processing object handler is undefined while processing files deployment config.", Logger.LogLevel.Error);
            processingPromisses.push(Promise.reject(undefined));
        }

        return Promise.all(processingPromisses);
    }

    private processDeploymentConfigNodesParallel(processingHandler: ISPObjectHandler, deploymentConfigNodes: Array<any>, dependentPromise: Promise<any>): Promise<any> {
        let processingPromisses: Array<Promise<any>> = [Promise.resolve()];

        if (processingHandler && deploymentConfigNodes && deploymentConfigNodes instanceof Array && deploymentConfigNodes.length > 0) {
            deploymentConfigNodes.forEach((processingConfig, proecssingIndex, array) => {
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
            processingPromise = deploymentConfigNodes.reduce((previousPromise, processingConfig, proecssingIndex, array) => {
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
        if (this._siteDeploymentConfig.User.proxyUrl) {
            NodeHttpProxy.url = url.parse(this._siteDeploymentConfig.User.proxyUrl);
            NodeHttpProxy.activate();
        }
    }

    private setupPnPJs(): void {
        let userConfig = this._siteDeploymentConfig.User;
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
                    siteUrl: this._siteDeploymentConfig.Site.Url,
                    username: userAndDommain[1],
                    workstation: userConfig.workstation,
                },
            };
        } else if (String(userConfig.authtype).toLowerCase() === AuthenticationType.Basic.toLowerCase()) {
            pnpConfig = {
                nodeHttpBasicClientOptions: {
                    password: userConfig.password,
                    siteUrl: this._siteDeploymentConfig.Site.Url,
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
