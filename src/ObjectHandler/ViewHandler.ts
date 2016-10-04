import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { List } from "@agileis/sp-pnp-js/lib/sharepoint/rest/lists";
import { View } from "@agileis/sp-pnp-js/lib/sharepoint/rest/views";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IView } from "../Interfaces/Types/IView";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { ControlOption } from "../Constants/ControlOption";
import { Util } from "../Util/Util";

export class ViewHandler implements ISPObjectHandler {
    public execute(viewConfig: IView, parentPromise: Promise<IPromiseResult<List>>): Promise<IPromiseResult<void | View>> {
        return new Promise<IPromiseResult<void | View>>((resolve, reject) => {
            parentPromise.then((promiseResult) => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, viewConfig.Title,
                        `View handler parent promise value result is null or undefined for the view with the title '${viewConfig.Title}'!`);
                } else {
                    let list = promiseResult.value;
                    this.processingViewConfig(viewConfig, list)
                        .then((viewProsssingResult) => { resolve(viewProsssingResult); })
                        .catch((error) => {
                            Util.Retry(error, viewConfig.Title,
                                () => {
                                    return this.processingViewConfig(viewConfig, list);
                                });
                        });
                }
            });
        });
    }

    private processingViewConfig(viewConfig: IView, list: List): Promise<IPromiseResult<void | View>> {
        return new Promise<IPromiseResult<void | View>>((resolve, reject) => {
            let processingText = viewConfig.ControlOption === ControlOption.Add || viewConfig.ControlOption === undefined || viewConfig.ControlOption === ""
                ? "Add" : viewConfig.ControlOption;
            Logger.write(`Processing '${processingText}' view: '${viewConfig.Title}.`, Logger.LogLevel.Info);

            list.views.filter(`Title eq '${viewConfig.Title}'`).select("Id").get()
                .then((viewRequestResults) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void | View>> = undefined;

                    if (viewRequestResults && viewRequestResults.length === 1) {
                        Logger.write(`Found view with title: '${viewConfig.Title}'`);
                        let view = list.views.getByTitle(viewConfig.Title);
                        switch (viewConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateView(viewConfig, view);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteView(viewConfig, view);
                                break;
                            default:
                                this.addViewFields(viewConfig, view)
                                    .then(() => {
                                        Util.Resolve<View>(resolve, viewConfig.Title, `Added viewfields to view with the title '${viewConfig.Title}'.`, view);
                                        rejectOrResolved = true;
                                    })
                                    .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while adding viewfields in the view with the title '${viewConfig.Title}': ` + error); });
                                break;
                        }
                    } else {
                        switch (viewConfig.ControlOption) {
                            case ControlOption.Delete:
                                Util.Resolve<void>(reject, viewConfig.Title, `View with the title '${viewConfig.Title}' does not have to be deleted, because it does not exist.`);
                                rejectOrResolved = true;
                                break;
                            case ControlOption.Update:
                                viewConfig.ControlOption = ControlOption.Add;
                            default:
                                processingPromise = this.addView(viewConfig, list);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((viewProsssingResult) => { resolve(viewProsssingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write("View handler processing promise is undefined!", Logger.LogLevel.Error);
                    }
                })
                .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while requesting view with the title '${viewConfig.Title}': ` + error); });
        });
    }

    private addView(viewConfig: IView, list: List): Promise<IPromiseResult<View>> {
        return new Promise<IPromiseResult<View>>((resolve, reject) => {
            let properties = this.createProperties(viewConfig);
            list.views.add(viewConfig.Title, viewConfig.PersonalView, properties)
                .then((viewAddResult) => {
                    this.addViewFields(viewConfig, viewAddResult.view)
                        .then(() => {
                            Util.Resolve<View>(resolve, viewConfig.Title, `Added view: '${viewConfig.Title}' and added all Viewfields.`, viewAddResult.view);
                        })
                        .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while adding Viewfields in the view with the title '${viewConfig.Title}': ` + error); });
                })
                .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while adding view with the title '${viewConfig.Title}': ` + error); });
        });
    }

    private updateView(viewConfig: IView, view: View): Promise<IPromiseResult<View>> {
        return new Promise<IPromiseResult<View>>((resolve, reject) => {
            let properties = this.createProperties(viewConfig);
            view.update(properties)
                .then((viewUpdateResult) => {
                    this.addViewFields(viewConfig, viewUpdateResult.view)
                        .then(() => {
                            Util.Resolve<View>(resolve, viewConfig.Title, `Updated view: '${viewConfig.Title}' and added all Viewfields.`, viewUpdateResult.view);
                        })
                        .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while adding Viewfields in the view with the title '${viewConfig.Title}': ` + error); });
                })
                .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while updating view with the title '${viewConfig.Title}': ` + error); });
        });
    }

    private deleteView(viewConfig: IView, view: View): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            view.delete()
                .then(() => { Util.Resolve<void>(resolve, viewConfig.Title, `Deleted view: '${viewConfig.Title}'.`); })
                .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while deleting view with the title '${viewConfig.Title}': ` + error); });
        });
    }

    private addViewFields(viewConfig: IView, view: View): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let viewUrl = view.toUrl();
            let listUrlParts = viewUrl.split("'");
            Logger.write(`Updating all viewfields from view with the title '${viewConfig.Title}' on the list '${listUrlParts[1]}'. `, Logger.LogLevel.Verbose);
            let context = SP.ClientContext.get_current();
            let spView: SP.View = context.get_web().get_lists().getByTitle(listUrlParts[1]).get_views().getByTitle(viewConfig.Title);
            let viewFieldCollection = spView.get_viewFields();
            viewFieldCollection.removeAll();
            for (let viewField of viewConfig.ViewFields) {
                let viewFieldName = viewField.InternalName;
                if (viewField.IsDependentLookup) {
                    viewFieldName = `${viewField.LookupListTitle}_${viewField.InternalName}`.substr(0, 32);
                }
                viewFieldCollection.add(viewFieldName);
            }

            spView.update();
            context.executeQueryAsync(
                (sender, args) => {
                    resolve();
                },
                (sender, args) => {
                    reject(args.get_message());
                });
        });
    }


    private createProperties(viewConfig: IView) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(viewConfig);
        let parsedObject: IView = JSON.parse(stringifiedObject);
        switch (viewConfig.ControlOption) {
            case ControlOption.Update:
                delete parsedObject.ControlOption;
                delete parsedObject.PersonalView;
                delete parsedObject.ViewFields;
                break;
            default:
                delete parsedObject.ControlOption;
                delete parsedObject.Title;
                delete parsedObject.PersonalView;
                delete parsedObject.ViewFields;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}
