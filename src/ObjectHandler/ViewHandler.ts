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
                        .catch((error) => { reject(error); });
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
                    let processingPromise: Promise<IPromiseResult<void | View>>  = undefined;

                    if (viewRequestResults && viewRequestResults.length === 1) {
                        let view = list.views.getByTitle(viewConfig.Title);
                        switch (viewConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateView(viewConfig, view);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteView(viewConfig, view);
                                break;
                            default:
                                Util.Resolve<View>(resolve, viewConfig.Title, `Added view with the title '${viewConfig.Title}', beacause it already exists.`, view);
                                rejectOrResolved = true;
                                break;
                        }
                    } else {
                        switch (viewConfig.ControlOption) {
                            case ControlOption.Delete:
                                Util.Resolve<void>(reject, viewConfig.Title, `Deleted view with title '${viewConfig.Title}', because it does not exists.`);
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
                    viewAddResult.view.fields.removeAll()
                        .then(() => { Util.Resolve<View>(resolve, viewConfig.Title, `Added view: '${viewConfig.Title}' and removed all default fields.`, viewAddResult.view); })
                        .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while deleting all fields in the view with the title '${viewConfig.Title}': ` + error); });
                })
                .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while adding view with the title '${viewConfig.Title}': ` + error); });
        });
    }

    private updateView(viewConfig: IView, view: View): Promise<IPromiseResult<View>> {
        return new Promise<IPromiseResult<View>>((resolve, reject) => {
            let properties = this.createProperties(viewConfig);
            view.update(properties)
                .then((viewUpdateResult) => {
                    viewUpdateResult.view.fields.removeAll()
                        .then(() => { Util.Resolve<View>(resolve, viewConfig.Title, `Updated view: '${viewConfig.Title}'.`, viewUpdateResult.view); })
                        .catch((error) => { Util.Reject<void>(reject, viewConfig.Title, `Error while deleting all fields in the view with the title '${viewConfig.Title}': ` + error); });
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
