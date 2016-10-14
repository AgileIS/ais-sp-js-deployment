import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ISPObjectHandler } from "../interfaces/objectHandler/iSpObjectHandler";
import { IPromiseResult } from "../interfaces/iPromiseResult";
import { INavigation } from "../interfaces/types/iNavigation";
import { INavigationNode } from "../interfaces/types/iNavigationNode";
import { Util } from "../util/util";

export class NavigationHandler implements ISPObjectHandler {
    private handlerName = "NavigationHandler";
    public execute(navigationConfig: INavigation, parentPromise: Promise<IPromiseResult<Web>>): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            parentPromise.then(promiseResult => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, this.handlerName,
                        `Navigation handler parent promise value result is null or undefined !`);
                } else {
                    let context = SP.ClientContext.get_current();
                    Util.tryToProcess("Navigation", () => { return this.processingQuicklaunchNavigationConfig(navigationConfig, context); }, this.handlerName)
                        .then(navigationProcessingResult => { resolve(navigationProcessingResult); })
                        .catch(error => { reject(error); });
                }
            });
        });
    }

    private processingQuicklaunchNavigationConfig(navigationConfig: INavigation, clientContext: SP.ClientContext): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            Logger.write(`${this.handlerName} - Processing quicklaunch navigation nodes.`, Logger.LogLevel.Info);

            let web = clientContext.get_web();
            let navigation = web.get_navigation();
            let quicklaunch = navigation.get_quickLaunch();

            let useSharedNavigation = navigationConfig.UseShared === true ? navigationConfig.UseShared : false;
            Logger.write(`${this.handlerName} - Set use shared navigation to ${useSharedNavigation}.`, Logger.LogLevel.Info);
            navigation.set_useShared(useSharedNavigation);

            clientContext.load(quicklaunch, "Include(Title,Url,IsExternal,Children)");
            clientContext.executeQueryAsync(
                (sender, args) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<SP.NavigationNodeCollection>> = undefined;

                    if (navigationConfig.ReCreateQuicklaunch) {
                        processingPromise = this.recreatingNavigationNodes(quicklaunch, navigationConfig.QuickLaunch);
                    } else {
                        processingPromise = this.updateNavigationNodeCollection(quicklaunch, navigationConfig.QuickLaunch);
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((quicklaunchProcessingResult) => { resolve(); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write("${this.handlerName} - Processing promise is undefined!", Logger.LogLevel.Error);
                    }
                },
                (sender, args) => {
                    Util.Reject(reject, this.handlerName, `Error while requesting quicklaunch node collection': `
                            + `${Util.getErrorMessageFromQuery(args.get_message(),args.get_stackTrace())}`);
                }
            );
        });
    }

    private recreatingNavigationNodes(navigationNodeCollection: SP.NavigationNodeCollection, navigatioNodes: Array<INavigationNode>): Promise<IPromiseResult<SP.NavigationNodeCollection>> {
        return new Promise<IPromiseResult<SP.NavigationNodeCollection>>((resolve, reject) => {
            Logger.write(`${this.handlerName} - Recreating quicklaunch.`, Logger.LogLevel.Info);

            this.clearNavigationNodeCollection(navigationNodeCollection);
            this.addNavNodesToNavCollection(navigationNodeCollection, navigatioNodes);

            navigationNodeCollection.get_context().executeQueryAsync(
                (sender2, args2) => {
                    Util.Resolve<SP.NavigationNodeCollection>(resolve, this.handlerName, `Recreated quicklaunch.`, navigationNodeCollection);
                },
                (sender2, args2) => {
                    Util.Reject<void>(reject, this.handlerName, `Error while recreating quicklaunch: ${args2.get_message()} '\n' ${args2.get_stackTrace()}`);
                }
            );
        });
    }

    private clearNavigationNodeCollection(navigationNodeCollection: SP.NavigationNodeCollection): void {
        Logger.write(`${this.handlerName} - Clearing navigation node collection.`, Logger.LogLevel.Info);

        let nodeEnumurator = navigationNodeCollection.getEnumerator();
        let toDeleteNodes: Array<SP.NavigationNode> = new Array<SP.NavigationNode>();
        while (nodeEnumurator.moveNext()) {
            toDeleteNodes.push(nodeEnumurator.get_current());
        }

        toDeleteNodes.forEach(
            (node, index, array) => {
                node.deleteObject();
            }
        );
    }

    private addNavNodesToNavCollection(navigationNodeCollection: SP.NavigationNodeCollection, navigatioNodes: Array<INavigationNode>): void {
        Logger.write(`${this.handlerName} - Adding navigation nodes to navigation node collection.`, Logger.LogLevel.Info);

        if (navigatioNodes) {
            navigatioNodes.forEach(
                (nodeConfig, index, array) => {
                    if (nodeConfig.Title && nodeConfig.Url) {
                        let nodeCreationInfo = new SP.NavigationNodeCreationInformation();
                        nodeCreationInfo.set_title(nodeConfig.Title);
                        nodeCreationInfo.set_url(nodeConfig.Url);
                        let isExternal = nodeConfig.IsExternal === true;
                        nodeCreationInfo.set_isExternal(isExternal);
                        nodeCreationInfo.set_asLastNode(true);

                        let navNode = navigationNodeCollection.add(nodeCreationInfo);
                        Logger.write(`${this.handlerName} - Added navigation node: ${nodeConfig.Title} - ${nodeConfig.Url}.`, Logger.LogLevel.Info);

                        if (nodeConfig.Children) {
                            this.addNavNodesToNavCollection(navNode.get_children(), nodeConfig.Children);
                        }
                    } else {
                        Logger.write(`${this.handlerName} - QuickLaunch navigation node ${index} missing title or/and url.`, Logger.LogLevel.Error);
                    }
                }
            );
        }
    }

    private getNavigationNodeByTitle(title: string, navigationNodeCollection: SP.NavigationNodeCollection): SP.NavigationNode {
        let navigationNode: SP.NavigationNode = undefined;
        if (navigationNodeCollection) {
            let nodeEnumurator = navigationNodeCollection.getEnumerator();
            while (nodeEnumurator.moveNext()) {
                let currentNode = nodeEnumurator.get_current();
                if (currentNode.get_title() === title) {
                    navigationNode = currentNode;
                    break;
                }

                let children = currentNode.get_children();
                if (children) {
                    let foundNodeInChilds = this.getNavigationNodeByTitle(title, children);
                    if (foundNodeInChilds) {
                        navigationNode = foundNodeInChilds;
                        break;
                    }
                }
            }
        }

        return navigationNode;
    }

    private updateNavigationNodeCollection(navigationNodeCollection: SP.NavigationNodeCollection, navigationNodes: Array<INavigationNode>): Promise<IPromiseResult<SP.NavigationNodeCollection>> {
        return new Promise<IPromiseResult<SP.NavigationNodeCollection>>((resolve, reject) => {
            Logger.write(`${this.handlerName} - Updating quicklaunch.`, Logger.LogLevel.Info);
            Util.Resolve<SP.NavigationNodeCollection>(resolve, this.handlerName, "Updated quicklaunch.", navigationNodeCollection);
            /* todo:
                        if (navigatioNodes) {
                            navigatioNodes.forEach((nodeConfig, index, array) => {
                                let navNode = this.getNavigationNodeByTitle(nodeConfig.Title, navigationNodeCollection);
                                if (navNode) {
                                    switch (nodeConfig.ControlOption) {
                                        case ControlOption.UPDATE:
                                            navNode
                                            break;
                                        case ControlOption.DELETE:
                                            navNode.deleteObject();
                                            break;
                                        default:
                                            Resolve(resolve, `View with the title '${viewConfig.Title}' already exists`, viewConfig.Title, view);
                                            break;
                                    }
                                } else {
                                    switch (nodeConfig.ControlOption) {
                                        case ControlOption.DELETE:
                                            Resolve(resolve, `Deleted quicklaunch navigation node with title '${nodeConfig.Title}'.`, "Navigation > Quicklaunch");
                                            break;
                                        case ControlOption.UPDATE:
            
                                        default:
                                            processingPromise = this.addView(viewConfig, parentList);
                                            break;
                                    }
                                }
                            });
                        } else {
                            Resolve(resolve, "Updated quicklaunch", "Navigation > Quicklaunch");
                        }*/
        });
    }
}
