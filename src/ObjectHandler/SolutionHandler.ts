import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { ISolution } from "../Interfaces/Types/ISolution";
import { File } from "@agileis/sp-pnp-js/lib/sharepoint/rest/files";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { ControlOption } from "../Constants/ControlOption";
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
                            .then(() => { resolve(); })
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
            let listRootFolder = clientContext.get_web().get_lists().getByTitle(solutionConfig.Library).get_rootFolder();
            clientContext.load(listRootFolder);
            clientContext.executeQueryAsync(
                (sender, args) => {
                    let processingPromise: Promise<IPromiseResult<void>> = undefined;
                    let packageInfo = new SP.Publishing.DesignPackageInfo();
                    packageInfo.set_packageGuid(SP.Guid.newGuid());
                    packageInfo.set_majorVersion(solutionConfig.MajorVersion);
                    packageInfo.set_minorVersion(solutionConfig.MinorVersion);
                    packageInfo.set_packageName(solutionConfig.Title);
                    let fileRelativeUrl = listRootFolder.get_serverRelativeUrl() + "/" + solutionConfig.Src + solutionConfig.FileName;
                    switch (solutionConfig.ControlOption) {
                        case ControlOption.Delete:
                            processingPromise = this.uninstallSolution(solutionConfig, clientContext, packageInfo);
                            break;
                        case ControlOption.Update:
                            this.uninstallSolution(solutionConfig, clientContext, packageInfo)
                                .then(() => { processingPromise = this.installSolution(solutionConfig, clientContext, packageInfo, fileRelativeUrl); })
                                .catch((error) => {
                                    processingPromise = Promise.reject(error);
                                });
                        default:
                            processingPromise = this.installSolution(solutionConfig, clientContext, packageInfo, fileRelativeUrl);
                            break;
                    }
                    processingPromise
                        .then((solutionProcessingResult) => { resolve(solutionProcessingResult); })
                        .catch((error) => { reject(error); });
                },
                (sender, args) => {
                    Util.Reject<void>(reject, solutionConfig.Title,
                        `Error while requesting Solution with the title '${solutionConfig.Title}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                }
            );
        });
    }

    private installSolution(solutionConfig: ISolution, clientContext: SP.ClientContext, packageInfo: SP.Publishing.DesignPackageInfo, filerelativeurl: string): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            SP.Publishing.DesignPackage.install(clientContext, clientContext.get_site(), packageInfo, filerelativeurl);
            clientContext.executeQueryAsync(
                (sender, args) => {
                    this.removeSolutionFile(solutionConfig, clientContext, filerelativeurl)
                        .then(() => { Util.Resolve<void>(resolve, solutionConfig.Title, `Activated Solution with title : '${solutionConfig.Title}'.`); })
                        .catch((error) => {
                            Util.Reject<void>(reject, solutionConfig.Title,
                                `Error while deleting Solution File with the title '${solutionConfig.Title}': ${Util.getErrorMessage(error)}`);
                        });
                },
                (sender, args) => {
                    Util.Reject<void>(reject, solutionConfig.Title,
                        `Error while installing Solution with the title '${solutionConfig.Title}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                });
        });
    }

    private uninstallSolution(solutionConfig: ISolution, clientContext: SP.ClientContext, packageInfo: SP.Publishing.DesignPackageInfo): Promise<IPromiseResult<void>> {
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

    private removeSolutionFile(solutionConfig: ISolution, clientContext: SP.ClientContext, fileRelativeUrl: string) {
        return new Promise<boolean>((resolve, reject) => {
            let item = clientContext.get_web().getFileByServerRelativeUrl(fileRelativeUrl);
            if (!item.get_serverObjectIsNull) {
                item.deleteObject();
                clientContext.executeQueryAsync(
                    (sender, args) => {
                        resolve();
                    },
                    (sender, args) => {
                        Util.Reject<void>(reject, solutionConfig.Title,
                            `Error while deleting Solution with the title '${solutionConfig.Title}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                    });
            } else {
                Util.Reject<void>(reject, solutionConfig.Title,
                    `Error while deleting Solutionfile '${solutionConfig.FileName}'' - file not found`);
            }

        });
    }
}
