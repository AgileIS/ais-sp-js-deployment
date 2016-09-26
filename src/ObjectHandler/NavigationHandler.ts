/// <reference path="../../typings/index.d.ts" />

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
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((quicklaunchProcessingResult) => {
                                resolve(quicklaunchProcessingResult);
                            })
                            .catch((error) => { reject(error); });
                    } else {
                        resolve();
                    }
                },
                (sender, args) => {
                    Reject(reject, `Error while requesting quicklaunch node collection': ${args.get_message()} '\n' ${args.get_stackTrace()}`, "Navigation > Quicklaunch");
                }
            );
        });
    }

    private recreatingQuicklaunch(quicklaunch: SP.NavigationNodeCollection, navigatioNodes: Array<INavigationNode>) {
        return new Promise<void>((resolve, reject) => {
            Logger.write("Recreating quicklaunch", Logger.LogLevel.Info);
            this.clearNavigationNodeCollection(quicklaunch);
            this.addNodesToQuickLaunch(quicklaunch, navigatioNodes);

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

    private clearNavigationNodeCollection(quicklaunch: SP.NavigationNodeCollection): void {
        let nodeEnumurator = quicklaunch.getEnumerator();
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

    private addNodesToQuickLaunch(quicklaunch: SP.NavigationNodeCollection, navigatioNodes: Array<INavigationNode>): void {
        navigatioNodes.forEach(
            (node, index, array) => {
                if (node.Title && node.Url) {
                    let nodeCreationInfo = new SP.NavigationNodeCreationInformation();
                    nodeCreationInfo.set_title(node.Title);
                    nodeCreationInfo.set_url(node.Url);
                    let IsExternal = node.IsExternal === true ? node.IsExternal : false;
                    nodeCreationInfo.set_isExternal(IsExternal);
                    nodeCreationInfo.set_asLastNode(true);
                    quicklaunch.add(nodeCreationInfo); // childnodes
                } else {
                    Logger.write(`QuickLaunch navigation node ${index} missing title or/and url`, Logger.LogLevel.Error);
                }
            }
        );
    }
}
