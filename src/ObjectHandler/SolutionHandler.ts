import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { ISolution } from "../Interfaces/Types/ISolution";
import { File } from "@agileis/sp-pnp-js/lib/sharepoint/rest/files";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { Util } from "../Util/Util";

export class SolutionHandler implements ISPObjectHandler {
    public execute(solutionConfig: ISolution, parentPromise: Promise<IPromiseResult<File>>): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            parentPromise.then((promiseResult) => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, solutionConfig.Title,
                        `Solution handler parent promise value result is null or undefined for the solution with the Title '${solutionConfig.Title}'!`);
                } else {
                    if (solutionConfig.Title) {
                        let context = SP.ClientContext.get_current();
                        this.processingSolutionConfig(solutionConfig, context)
                            .then((featureProsssingResult) => { resolve(featureProsssingResult); })
                            .catch((error) => {
                                Util.Retry(error, solutionConfig.Title,
                                    () => {
                                        return this.processingSolutionConfig(solutionConfig, context);
                                    });
                            });
                    } else {
                        Util.Reject<void>(reject, "Unknow Solution", `Error while processing Solution: Solution Title is undefined.`);
                    }
                }
            });
        });
    }

    private processingSolutionConfig(solutionConfig: ISolution, clientContext: SP.ClientContext): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            Logger.write(`Processing Solution: '${solutionConfig.Title}'.`, Logger.LogLevel.Info);
            let solutionGallery = clientContext.get_web().get_lists().getByTitle(solutionConfig.Library);
            let solutionGalRootFolder = solutionGallery.get_rootFolder();
            let qry = new SP.CamlQuery();
            qry.set_viewXml(`<Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='File'>${solutionConfig.Title}</Value></Eq></Where>`);
            let itemColl = solutionGallery.getItems(qry);
            clientContext.load(itemColl);
            clientContext.load(solutionGalRootFolder);
            clientContext.executeQueryAsync(
                (sender, args) => {
                    if (itemColl.get_count() === 1) {
                        let processingPromise: Promise<IPromiseResult<void>> = undefined;
                        let packageInfo = new SP.Publishing.DesignPackageInfo();
                        packageInfo.set_packageGuid(SP.Guid.newGuid());
                        packageInfo.set_majorVersion(solutionConfig.MajorVersion);
                        packageInfo.set_minorVersion(solutionConfig.MinorVersion);
                        packageInfo.set_packageName(solutionConfig.Title);
                        let fileRelativeUrl = solutionGalRootFolder.get_serverRelativeUrl() + `/${solutionConfig.Title}`;
                        switch (solutionConfig.Deactivate) {
                            case false:
                                processingPromise = this.activateSolution(solutionConfig, clientContext, packageInfo, fileRelativeUrl);
                                break;
                            default:
                                processingPromise = this.deactivateSolution(solutionConfig, clientContext, packageInfo);
                                break;
                        }
                        processingPromise
                            .then(() => { resolve(); })
                            .catch((error) => { reject(error); });
                    } else if (itemColl.get_count() === 0) {
                        Util.Reject<void>(reject, solutionConfig.Title, `No Solution with the title '${solutionConfig.Title}' found`);
                    } else if (itemColl.get_count() > 1) {
                        Util.Reject<void>(reject, solutionConfig.Title, `More than one Solution with the title '${solutionConfig.Title}' found`);
                    }
                },
                (sender, args) => {
                    Util.Reject<void>(reject, solutionConfig.Title,
                        `Error while requesting Solution with the title '${solutionConfig.Title}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                }
            );
        });
    }

    private activateSolution(solutionConfig: ISolution, clientContext: SP.ClientContext, packageInfo: SP.Publishing.DesignPackageInfo, filerelativeurl: string) {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            SP.Publishing.DesignPackage.install(clientContext, clientContext.get_site(), packageInfo, filerelativeurl);
            clientContext.executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<void>(resolve, solutionConfig.Title, `Activated Solution with title : '${solutionConfig.Title}'.`);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, solutionConfig.Title,
                        `Error while activating Solution with the title '${solutionConfig.Title}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                });
        });
    }

    private deactivateSolution(solutionConfig: ISolution, clientContext: SP.ClientContext, packageInfo: SP.Publishing.DesignPackageInfo) {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            SP.Publishing.DesignPackage.unInstall(clientContext, clientContext.get_site(), packageInfo);
            clientContext.executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<void>(resolve, solutionConfig.Title, `Deactivated Solution with title : '${solutionConfig.Title}'.`);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, solutionConfig.Title,
                        `Error while deactivating Solution with the title '${solutionConfig.Title}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                });
        });
    }

}
