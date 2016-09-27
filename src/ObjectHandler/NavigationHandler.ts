import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { INavigation } from "../Interfaces/Types/INavigation";
import { INavigationNode } from "../Interfaces/Types/INavigationNode";
import { ControlOption } from "../Constants/ControlOption";
import { Util } from "../Util/Util";

export class NavigationHandler implements ISPObjectHandler {
    public execute(navigationConfig: INavigation, parentPromise: Promise<IPromiseResult<Web>>): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            parentPromise.then(promiseResult => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, "Navigation > Quicklaunch",
                        `Navigation handler parent promise value result is null or undefined !`);
                } else {
                    let context = SP.ClientContext.get_current();
                    this.processingQuicklaunchNavigationConfig(navigationConfig, context)
                        .then((NavigationProsssingResult) => { resolve(NavigationProsssingResult); })
                        .catch((error) => {
                            Util.Retry(error, "Navigation > Quicklaunch",
                                () => {
                                    return this.processingQuicklaunchNavigationConfig(navigationConfig, context);
                                });
                        });
                }
            });
        });
    }

    private processingQuicklaunchNavigationConfig(navigationConfig: INavigation, clientContext: SP.ClientContext): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            Logger.write(`Processing quicklaunch navigation nodes.`, Logger.LogLevel.Info);

            let web = clientContext.get_web();
            let navigation = web.get_navigation();
            let quicklaunch = navigation.get_quickLaunch();

            let useSharedNavigation = navigationConfig.UseShared === true ? navigationConfig.UseShared : false;
            Logger.write(`Set use shared navigation to ${useSharedNavigation}.`, Logger.LogLevel.Info);
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
                        Logger.write("Navigation handler processing promise is undefined!", Logger.LogLevel.Error);
                    }
                },
                (sender, args) => {
                    Util.Reject(reject, "Navigation > Quicklaunch", `Error while requesting quicklaunch node collection': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                }
            );
        });
    }

    private recreatingNavigationNodes(navigationNodeCollection: SP.NavigationNodeCollection, navigatioNodes: Array<INavigationNode>): Promise<IPromiseResult<SP.NavigationNodeCollection>> {
        return new Promise<IPromiseResult<SP.NavigationNodeCollection>>((resolve, reject) => {
            Logger.write("Recreating quicklaunch.", Logger.LogLevel.Info);

            this.clearNavigationNodeCollection(navigationNodeCollection);
            this.addNavNodesToNavCollection(navigationNodeCollection, navigatioNodes);

            navigationNodeCollection.get_context().executeQueryAsync(
                (sender2, args2) => {
                    Util.Resolve<SP.NavigationNodeCollection>(resolve, "Navigation > Quicklaunch", `Recreated quicklaunch.`, navigationNodeCollection);
                },
                (sender2, args2) => {
                    Util.Reject<void>(reject, "Navigation > Quicklaunch", `Error while recreating quicklaunch: ${args2.get_message()} '\n' ${args2.get_stackTrace()}`);
                }
            );
        });
    }

    private clearNavigationNodeCollection(navigationNodeCollection: SP.NavigationNodeCollection): void {
        Logger.write("Clearing navigation node collection.", Logger.LogLevel.Info);

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
        Logger.write("Adding navigation nodes to navigation node collection.", Logger.LogLevel.Info);

        if (navigatioNodes) {
            navigatioNodes.forEach(
                (nodeConfig, index, array) => {
                    if (nodeConfig.Title && nodeConfig.Url) {
                        let nodeCreationInfo = new SP.NavigationNodeCreationInformation();
                        nodeCreationInfo.set_title(nodeConfig.Title);
                        nodeCreationInfo.set_url(nodeConfig.Url);
                        let IsExternal = nodeConfig.IsExternal === true ? nodeConfig.IsExternal : false;
                        nodeCreationInfo.set_isExternal(IsExternal);
                        nodeCreationInfo.set_asLastNode(true);

                        let navNode = navigationNodeCollection.add(nodeCreationInfo);
                        Logger.write(`Added navigation node: ${nodeConfig.Title} - ${nodeConfig.Url}.`, Logger.LogLevel.Info);

                        if (nodeConfig.Children) {
                            this.addNavNodesToNavCollection(navNode.get_children(), nodeConfig.Children);
                        }
                    } else {
                        Logger.write(`QuickLaunch navigation node ${index} missing title or/and url.`, Logger.LogLevel.Error);
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
            Logger.write("Updating quicklaunch.", Logger.LogLevel.Info);
            Util.Resolve<SP.NavigationNodeCollection>(resolve, "Navigation > Quicklaunch", "Updated quicklaunch.", navigationNodeCollection);
            /* todo:
                        if (navigatioNodes) {
                            navigatioNodes.forEach((nodeConfig, index, array) => {
                                let navNode = this.getNavigationNodeByTitle(nodeConfig.Title, navigationNodeCollection);
                                if (navNode) {
                                    switch (nodeConfig.ControlOption) {
                                        case ControlOption.Update:
                                            navNode
                                            break;
                                        case ControlOption.Delete:
                                            navNode.deleteObject();
                                            break;
                                        default:
                                            Resolve(resolve, `View with the title '${viewConfig.Title}' already exists`, viewConfig.Title, view);
                                            break;
                                    }
                                } else {
                                    switch (nodeConfig.ControlOption) {
                                        case ControlOption.Delete:
                                            Resolve(resolve, `Deleted quicklaunch navigation node with title '${nodeConfig.Title}'.`, "Navigation > Quicklaunch");
                                            break;
                                        case ControlOption.Update:
            
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
