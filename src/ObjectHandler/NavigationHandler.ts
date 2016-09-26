import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { INavigation } from "../interface/Types/INavigation";
import { INavigationNode } from "../interface/Types/inavigationnode";
import { Resolve, Reject } from "../Util/Util";

export class NavigationHandler implements ISPObjectHandler {
    public execute(navigationConfig: INavigation, parentPromise: Promise<Web>): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            parentPromise.then(parentWeb => {
                this.processingQuicklaunchNavigationConfig(navigationConfig, parentWeb)
                    .then((NavigationProsssingResult) => { resolve(NavigationProsssingResult); })
                    .catch((error) => { reject(error); });
            });
        });
    }

    private processingQuicklaunchNavigationConfig(navigationConfig: INavigation, parentWeb: Web): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Logger.write(`Processing quicklaunch navigation nodes`, Logger.LogLevel.Info);

            let context = new SP.ClientContext(parentWeb.toUrl().split("/_")[0]);
            let web = context.get_web();
            let navigation = web.get_navigation();
            navigation.set_useShared(navigationConfig.UseShared === true ? navigationConfig.UseShared : false);
            let quicklaunch = navigation.get_quickLaunch();
            context.load(quicklaunch);
            context.executeQueryAsync(
                (sender, args) => {
                    let processingPromise: Promise<void> = undefined;

                    if (navigationConfig.ReCreateQuicklaunch) {
                        processingPromise = this.recreatingQuicklaunch(quicklaunch, navigationConfig.QuickLaunch);
                    }else{
                        processingPromise = this.insertNodesInQuicklaunch();
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((quicklaunchProcessingResult) => {
                                resolve(quicklaunchProcessingResult);
                            })
                            .catch((error) => { reject(error); });
                    } else {
                        Logger.write("Navigation handler processing promise is undefined!");
                    }
                },
                (sender, args) => {
                    Reject(reject, `Error while requesting quicklaunch node collection': ${args.get_message()} '\n' ${args.get_stackTrace()}`, "Navigation > Quicklaunch");
                }
            );
        });
    }

    private recreatingQuicklaunch(quicklaunch: SP.NavigationNodeCollection, navigatioNodes: Array<INavigationNode>): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Logger.write("Recreating quicklaunch", Logger.LogLevel.Info);

            this.clearNavigationNodeCollection(quicklaunch);
            this.addNavNodesToNavCollection(quicklaunch, navigatioNodes);

            quicklaunch.get_context().executeQueryAsync(
                (sender2, args2) => {
                    Resolve(resolve, `Recreated quicklaunch`, "Navigation > Quicklaunch");
                },
                (sender2, args2) => {
                    Reject(reject, `Error while recreating quicklaunch: ${args2.get_message()} '\n' ${args2.get_stackTrace()}`, "Navigation > Quicklaunch");
                }
            );
        });
    }

    private clearNavigationNodeCollection(nodeNavigationCollection: SP.NavigationNodeCollection): void {
        Logger.write("Clearing navigation node collection", Logger.LogLevel.Info);

        let nodeEnumurator = nodeNavigationCollection.getEnumerator();
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

    private addNavNodesToNavCollection(nodeNavigationCollection: SP.NavigationNodeCollection, navigatioNodes: Array<INavigationNode>): void {
        Logger.write("Adding navigation nodes to navigation node collection", Logger.LogLevel.Info);

        navigatioNodes.forEach(
            (nodeConfig, index, array) => {
                if (nodeConfig.Title && nodeConfig.Url) {
                    let nodeCreationInfo = new SP.NavigationNodeCreationInformation();
                    nodeCreationInfo.set_title(nodeConfig.Title);
                    nodeCreationInfo.set_url(nodeConfig.Url);
                    let IsExternal = nodeConfig.IsExternal === true ? nodeConfig.IsExternal : false;
                    nodeCreationInfo.set_isExternal(IsExternal);
                    nodeCreationInfo.set_asLastNode(true);

                    let navNode = nodeNavigationCollection.add(nodeCreationInfo);

                    if (nodeConfig.Children) {
                        this.addNavNodesToNavCollection(navNode.get_children(), nodeConfig.Children);
                    }
                } else {
                    Logger.write(`QuickLaunch navigation node ${index} missing title or/and url`, Logger.LogLevel.Error);
                }
            }
        );
    }

    private insertNodesInQuicklaunch(): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            resolve();
        });
    }
}
