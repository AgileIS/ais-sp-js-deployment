import { Logger } from "sp-pnp-js/lib/utils/logging";
import { List } from "sp-pnp-js/lib/sharepoint/rest/lists";
import { View } from "sp-pnp-js/lib/sharepoint/rest/views";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IView } from "../interface/Types/IView";
import { ControlOption } from "../Constants/ControlOption";
import { Reject, Resolve } from "../Util/Util";

export class ViewHandler implements ISPObjectHandler {
    public execute(viewConfig: IView, parentPromise: Promise<List>): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            parentPromise.then((parentList) => {
                this.ProcessingViewConfig(viewConfig, parentList).then((viewProsssingResult) => { resolve(viewProsssingResult); }).catch((error) => { reject(error); });
            });
        });
    }

    private ProcessingViewConfig(viewConfig: IView, parentList: List): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            Logger.write(`Processing ${viewConfig.ControlOption === ControlOption.Add || viewConfig.ControlOption === undefined ? "Add" : viewConfig.ControlOption} view: '${viewConfig.Title}'`, Logger.LogLevel.Info);
            parentList.views.filter(`Title eq '${viewConfig.Title}'`).select("Id").get().then((viewRequestResults) => {
                let processingPromise: Promise<View> = undefined;

                if (viewRequestResults && viewRequestResults.length === 1) {
                    let view = parentList.views.getByTitle(viewConfig.Title);
                    switch (viewConfig.ControlOption) {
                        case ControlOption.Update:
                            processingPromise = this.UpdateView(viewConfig, parentList, view);
                            break;
                        case ControlOption.Delete:
                            processingPromise = this.DeleteView(viewConfig, parentList, view);
                            break;
                        default:
                            Resolve(reject, `View with the title '${viewConfig.Title}' already exists`, viewConfig.Title, view);
                            break;
                    }
                } else {
                    switch (viewConfig.ControlOption) {
                        case ControlOption.Update:
                        case ControlOption.Delete:
                            Reject(reject, `View with title '${viewConfig.Title}' does not exists`, viewConfig.Title);
                            break;
                        default:
                            processingPromise = this.AddView(viewConfig, parentList);
                            break;
                    }
                }

                if (processingPromise) {
                    processingPromise.then((viewProsssingResult) => { resolve(viewProsssingResult); }).catch((error) => { reject(error); });
                }
            }).catch((error) => { Reject(reject, `Error while requesting view with the title '${viewConfig.Title}': ` + error, viewConfig.Title); });
        });
    }

    private AddView(viewConfig: IView, parentList: List): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            let properties = this.CreateProperties(viewConfig);
            parentList.views.add(viewConfig.Title, viewConfig.PersonalView, properties).then((viewAddResult) => {
                viewAddResult.view.fields.removeAll().then(() => {
                    Resolve(resolve, `Added view: '${viewConfig.Title}'`, viewConfig.Title, viewAddResult.view);
                }).catch((error) => { Reject(reject, `Error while deleting all fields in the view with the title '${viewConfig.Title}': ` + error, viewConfig.Title); });
            }).catch((error) => { Reject(reject, `Error while adding view with the title '${viewConfig.Title}': ` + error, viewConfig.Title); });
        });
    }

    private UpdateView(viewConfig: IView, parentList: List, view: View): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            let properties = this.CreateProperties(viewConfig);
            view.update(properties).then((viewUpdateResult) => {
                viewUpdateResult.view.fields.removeAll().then(() => {
                    Resolve(resolve, `Updated view: '${viewConfig.Title}'`, viewConfig.Title, viewUpdateResult.view);
                }).catch((error) => { Reject(reject, `Error while deleting all fields in the view with the title '${viewConfig.Title}': ` + error, viewConfig.Title); });
            }).catch((error) => { Reject(reject, `Error while updating view with the title '${viewConfig.Title}': ` + error, viewConfig.Title); });
        });
    }

    private DeleteView(viewConfig: IView, parentList: List, view: View): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            view.delete().then(() => {
                Resolve(resolve, `Deleted view: '${viewConfig.Title}'`, viewConfig.Title);
            }).catch((error) => { Reject(reject, `Error while deleting view with the title '${viewConfig.Title}': ` + error, viewConfig.Title); });
        });
    }

    private CreateProperties(viewConfig: IView) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(viewConfig);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (viewConfig.ControlOption) {
            case ControlOption.Update:
                delete parsedObject.ControlOption;
                delete parsedObject.PersonalView;
                delete parsedObject.ViewField;
                break;
            default:
                delete parsedObject.ControlOption;
                delete parsedObject.Title;
                delete parsedObject.PersonalView;
                delete parsedObject.ViewField;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}