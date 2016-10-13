import * as PnP from "@agileis/sp-pnp-js";
import { LibraryConfiguration } from "@agileis/sp-pnp-js/lib/configuration/pnplibconfig";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Folder } from "@agileis/sp-pnp-js/lib/sharepoint/rest/folders";
import { List } from "@agileis/sp-pnp-js/lib/sharepoint/rest/lists";
import { ISiteDeploymentConfig } from "./interfaces/config/iSiteDeploymentConfig";
import { ISPObjectHandler } from "./interfaces/objectHandler/iSpObjectHandler";
import { ISPObjectHandlerCollection } from "./interfaces/objectHandler/iSpObjectHandlerCollection";
import { IList } from "./interfaces/types/iList";
import { IFile } from "./interfaces/types/iFile";
import { IFolder } from "./interfaces/types/iFolder";
import { IPromiseResult } from "./interfaces/iPromiseResult";
import { SiteHandler } from "./objectHandler/siteHandler";
import { ListHandler } from "./objectHandler/listHandler";
import { ItemHandler } from "./objectHandler/itemHandler";
import { FileHandler } from "./objectHandler/fileHandler";
import { FieldHandler } from "./objectHandler/fieldHandler";
import { ViewHandler } from "./objectHandler/viewHandler";
import { SolutionHandler } from "./objectHandler/solutionHandler";
import { FeatureHandler } from "./objectHandler/featureHandler";
import { ContentTypeHandler } from "./objectHandler/contentTypeHandler";
import { NavigationHandler } from "./objectHandler/navigationHandler";
import { AuthenticationType } from "./constants/authenticationType";
import { NodeHttpProxy } from "./nodeHttpProxy";
import { NodeJsomHandler } from "./nodeJsomHandler";
import { Util } from "./util/util";
import * as url from "url";

export class DeploymentManager {
    private siteDeploymentConfig: ISiteDeploymentConfig;
    private deployDependencies: Array<Promise<any>> = new Array();
    private objectHandlers: ISPObjectHandlerCollection = {
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

    constructor(siteDeploymentConfig: ISiteDeploymentConfig) {
        if (siteDeploymentConfig.Site && siteDeploymentConfig.Site.Url) {
            let layoutsUrlPart = siteDeploymentConfig.Site.LayoutsHive ? `_layouts/${siteDeploymentConfig.Site.LayoutsHive}` : `_layouts/15`;
            this.siteDeploymentConfig = <ISiteDeploymentConfig>JSON.parse(
                Util.replaceUrlTokens(JSON.stringify(siteDeploymentConfig), Util.getRelativeUrl(siteDeploymentConfig.Site.Url), layoutsUrlPart));
            this.setupProxy();
            this.setupPnPJs();
            this.deployDependencies.push(NodeJsomHandler.initialize(siteDeploymentConfig));
        } else {
            throw new Error("Deployment config site or site url is undefined");
        }
    }

    public deploy(): Promise<void> {
        return Promise.all(this.deployDependencies).then(() => {
            return this.processDeploymentConfig()
                .then(() => {
                    Logger.write(`site collection '${this.siteDeploymentConfig.Site.Url}' processed`, Logger.LogLevel.Info);
                })
                .catch((error) => {
                    Logger.write(`Error occured while processing site collection '${this.siteDeploymentConfig.Site.Url}' - ` + Util.getErrorMessage(error), Logger.LogLevel.Error);
                });
        });
    }

    private processDeploymentConfig(): Promise<any> {
        let siteProcessingPromise = this.objectHandlers.Sites.execute(this.siteDeploymentConfig.Site, Promise.resolve());

        let nodeProcessingOrder: string[] = ["Features", "Fields", "ContentTypes", "Lists", "Navigation", "Files", "Solutions"];
        let existingSiteNodes = Object.keys(this.siteDeploymentConfig.Site);

        return nodeProcessingOrder.reduce((dependentPromise, processingKey, proecssingIndex, array): Promise<any> => {
            return dependentPromise
                .then(() => {
                    let processingConfig = (<any>this.siteDeploymentConfig.Site)[processingKey];
                    let processingHandler = this.objectHandlers[processingKey];
                    let processingPromise: Promise<any> = Promise.resolve();

                    if (existingSiteNodes.indexOf(processingKey) > -1 && processingHandler) {
                        if (processingKey === "Fields") {
                            processingPromise = this.processDeploymentConfigNodesParallel(processingHandler, processingConfig, siteProcessingPromise);
                        } else if (processingKey === "Features" || processingKey === "ContentTypes" || processingKey === "Solutions") {
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
                        return this.processDeploymentConfigNodesSequential(this.objectHandlers.Fields, listConfig.Fields, listPromise);
                    })
                    .then(() => {
                        return Promise.all([
                            this.processDeploymentConfigNodesParallel(this.objectHandlers.Views, listConfig.Views, listPromise),
                            this.processDeploymentConfigNodesParallel(this.objectHandlers.Items, listConfig.Items, listPromise),
                            this.processFilesDeploymentConfig(this.objectHandlers.Files, listConfig.Files, listPromise),
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
        if (this.siteDeploymentConfig.User.proxyUrl) {
            NodeHttpProxy.url = url.parse(this.siteDeploymentConfig.User.proxyUrl);
            NodeHttpProxy.activate();
        }
    }

    private setupPnPJs(): void {
        let userConfig = this.siteDeploymentConfig.User;
        Logger.write("Setup pnp-core-js", Logger.LogLevel.Info);
        Logger.write(`pnp-core-js authentication type: ${userConfig.authtype}`, Logger.LogLevel.Info);

        let pnpConfig: LibraryConfiguration;
        if (String(userConfig.authtype).toLowerCase() === AuthenticationType.NTLM.toLowerCase()) {
            let userAndDommain = userConfig.username.split("\\");
            if (!userConfig.workstation) {
                userConfig.workstation = "";
            }

            pnpConfig = {
                nodeHttpNtlmClientOptions: {
                    domain: userAndDommain[0],
                    password: userConfig.password,
                    siteUrl: this.siteDeploymentConfig.Site.Url,
                    username: userAndDommain[1],
                    workstation: userConfig.workstation,
                },
            };
        } else if (String(userConfig.authtype).toLowerCase() === AuthenticationType.BASIC.toLowerCase()) {
            pnpConfig = {
                nodeHttpBasicClientOptions: {
                    password: userConfig.password,
                    siteUrl: this.siteDeploymentConfig.Site.Url,
                    username: userConfig.username,
                },
            };
        } else {
            throw new Error(`Unsupported authentication type. Use ${AuthenticationType.NTLM} or ${AuthenticationType.BASIC} `);
        }

        if (pnpConfig) {
            PnP.setup(pnpConfig);
        }
    }
}
