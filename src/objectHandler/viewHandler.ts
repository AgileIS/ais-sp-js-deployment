import { Logger } from "ais-sp-pnp-js/lib/utils/logging";
import { List } from "ais-sp-pnp-js/lib/sharepoint/rest/lists";
import { View } from "ais-sp-pnp-js/lib/sharepoint/rest/views";
import { ISPObjectHandler } from "../interfaces/objectHandler/iSpObjectHandler";
import { IView } from "../interfaces/types/iView";
import { IPromiseResult } from "../interfaces/iPromiseResult";
import { ControlOption } from "../constants/controlOption";
import { Util } from "../util/util";

export class ViewHandler implements ISPObjectHandler {
    private handlerName = "ViewHandler";
    public execute(viewConfig: IView, parentPromise: Promise<IPromiseResult<List>>): Promise<IPromiseResult<void | View>> {
        return new Promise<IPromiseResult<void | View>>((resolve, reject) => {
            parentPromise.then((promiseResult) => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, this.handlerName,
                        `View handler parent promise value result is null or undefined for the view with the title '${viewConfig.Title}'!`);
                } else {
                    let list = promiseResult.value;
                    Util.tryToProcess(viewConfig.InternalName, () => { return this.processingViewConfig(viewConfig, list); }, this.handlerName)
                        .then(viewProcessingResult => { resolve(viewProcessingResult); })
                        .catch(error => { reject(error); });
                }
            });
        });
    }

    private processingViewConfig(viewConfig: IView, list: List): Promise<IPromiseResult<void | View>> {
        return new Promise<IPromiseResult<void | View>>((resolve, reject) => {
            let processingText = viewConfig.ControlOption === ControlOption.ADD || viewConfig.ControlOption === undefined || viewConfig.ControlOption === ""
                ? "Add" : viewConfig.ControlOption;
            Logger.write(`${this.handlerName} - Processing '${processingText}' view: '${viewConfig.Title}.`, Logger.LogLevel.Info);

            list.views.filter(`substringof('${viewConfig.InternalName}',ServerRelativeUrl) eq true`).select("Id", "Title").get()
                .then((viewRequestResults) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void | View>> = undefined;

                    if (viewRequestResults && viewRequestResults.length === 1) {
                        viewConfig.NewTitle = viewConfig.Title;
                        viewConfig.Title = viewRequestResults[0].Title;
                        Logger.write(`${this.handlerName} - Found view with title: '${viewConfig.Title}'`, Logger.LogLevel.Info);
                        let view = list.views.getByTitle(viewConfig.Title);
                        switch (viewConfig.ControlOption) {
                            case ControlOption.UPDATE:
                                processingPromise = this.updateView(viewConfig, view);
                                break;
                            case ControlOption.DELETE:
                                processingPromise = this.deleteView(viewConfig, view);
                                break;
                            default:
                                processingPromise = this.addViewFields(viewConfig, view);
                                rejectOrResolved = true;
                                break;
                        }
                    } else {
                        switch (viewConfig.ControlOption) {
                            case ControlOption.DELETE:
                                Util.Resolve<void>(reject, this.handlerName, `View with the title '${viewConfig.Title}' does not have to be deleted, because it does not exist.`);
                                rejectOrResolved = true;
                                break;
                            case ControlOption.UPDATE:
                                viewConfig.ControlOption = ControlOption.ADD;
                            default:
                                processingPromise = this.addView(viewConfig, list);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((viewProcessingResult) => { resolve(viewProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write(`${this.handlerName} - View handler processing promise is undefined!`, Logger.LogLevel.Error);
                    }
                })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while requesting view with the title '${viewConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private addView(viewConfig: IView, list: List): Promise<IPromiseResult<View>> {
        return new Promise<IPromiseResult<View>>((resolve, reject) => {
            let properties = this.createProperties(viewConfig);
            list.views.add(viewConfig.Title, viewConfig.PersonalView, properties)
                .then((viewAddResult) => {
                    this.addViewFields(viewConfig, viewAddResult.view)
                        .then(() => { Util.Resolve<View>(resolve, this.handlerName, `Added view: '${viewConfig.Title}' and added all Viewfields.`, viewAddResult.view); })
                        .catch((error) => {
                            Util.Reject<void>(reject, this.handlerName,
                                `Error while adding Viewfields in the view with the title '${viewConfig.Title}': ` + Util.getErrorMessage(error));
                        });
                })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while adding view with the title '${viewConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private updateView(viewConfig: IView, view: View): Promise<IPromiseResult<View>> {
        return new Promise<IPromiseResult<View>>((resolve, reject) => {
            let properties = this.createProperties(viewConfig);
            view.update(properties)
                .then((viewUpdateResult) => {
                    let viewTitle = this.getTitleFromConfig(viewConfig);
                    this.addViewFields(viewConfig, viewUpdateResult.view)
                        .then(() => {
                            Util.Resolve<View>(resolve, this.handlerName, `Updated view: '${viewTitle}' and added all Viewfields.`, viewUpdateResult.view);
                        })
                        .catch((error) => {
                            Util.Reject<void>(reject, this.handlerName,
                                `Error while adding Viewfields in the view with the title '${viewTitle}': ` + Util.getErrorMessage(error));
                        });
                })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while updating view with the title '${viewConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private deleteView(viewConfig: IView, view: View): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            view.delete()
                .then(() => { Util.Resolve<void>(resolve, this.handlerName, `Deleted view: '${viewConfig.Title}'.`); })
                .catch((error) => { Util.Reject<void>(reject, this.handlerName, `Error while deleting view with the title '${viewConfig.Title}': ` + Util.getErrorMessage(error)); });
        });
    }

    private addViewFields(viewConfig: IView, view: View): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let spView = undefined;
            let viewUrl = view.toUrl();
            let listUrlParts = viewUrl.split("'");
            let viewTitle = this.getTitleFromConfig(viewConfig);
            Logger.write(`${this.handlerName} - Updating all viewfields from view with the title '${viewConfig.Title}' on the list '${listUrlParts[1]}'. `, Logger.LogLevel.Verbose);
            let context = SP.ClientContext.get_current();
            if (listUrlParts[1].match(/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/)) {
                spView = context.get_web().get_lists().getById(listUrlParts[1]).get_views().getByTitle(viewTitle);
            } else {
                spView = context.get_web().get_lists().getByTitle(listUrlParts[1]).get_views().getByTitle(viewTitle);
            }

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
                    Util.Resolve<void>(resolve, this.handlerName, `Added viewfields to view with the title '${viewConfig.Title}'.`);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, this.handlerName, `Error while adding viewfields in the view with the title '${viewConfig.Title}': `
                            + `${Util.getErrorMessageFromQuery(args.get_message(),args.get_stackTrace())}`);
                });
        });
    }

    private getTitleFromConfig(viewConfig: IView): string {
        let viewTitle = viewConfig.Title;
        if (viewConfig.NewTitle && viewConfig.NewTitle !== viewConfig.Title && viewConfig.NewTitle.length > 0) {
            viewTitle = viewConfig.NewTitle;
        }
        return viewTitle;
    }

    private createProperties(viewConfig: IView) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(viewConfig);
        let parsedObject: IView = JSON.parse(stringifiedObject);
        switch (viewConfig.ControlOption) {
            case ControlOption.UPDATE:
                delete parsedObject.ControlOption;
                delete parsedObject.PersonalView;
                delete parsedObject.ViewFields;
                delete parsedObject.InternalName;
                delete parsedObject.NewTitle;
                parsedObject.Title = this.getTitleFromConfig(viewConfig);
                break;
            default:
                delete parsedObject.ControlOption;
                delete parsedObject.Title;
                delete parsedObject.NewTitle;
                delete parsedObject.InternalName;
                delete parsedObject.PersonalView;
                delete parsedObject.ViewFields;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}
