import { Logger } from "sp-pnp-js/lib/utils/logging";
import { List } from "sp-pnp-js/lib/sharepoint/rest/lists";
import { View } from "sp-pnp-js/lib/sharepoint/rest/views";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IView } from "../interface/Types/IView";
import { Reject, Resolve } from "../Util/Util";

export class ViewHandler implements ISPObjectHandler {
    public execute(viewConfig: IView, parentPromise: Promise<List>): Promise<View> {
        switch (viewConfig.ControlOption) {
            case "Update":
                return this.UpdateView(viewConfig, parentPromise);
            case "Delete":
                return this.DeleteView(viewConfig, parentPromise);
            default:
                return this.AddView(viewConfig, parentPromise);
        }
    }

    private AddView(viewConfig: IView, parentPromise: Promise<List>): Promise<View> {
        //todo: get view in execute
        return new Promise<View>((resolve, reject) => {
            parentPromise.then((parentInstance) => {
                Logger.write(`Adding view: '${viewConfig.Title}'`, Logger.LogLevel.Info);
                let view = parentInstance.views.getByTitle(viewConfig.Title);
                parentInstance.views.filter(`Title eq '${viewConfig.Title}'`).select("Id").get().then((result) => {
                    if (result.length === 0) {
                        let properties = this.CreateProperties(viewConfig);
                        parentInstance.views.add(viewConfig.Title, viewConfig.PersonalView, properties).then((result) => {
                            result.view.fields.removeAll().then(() => {
                                Resolve(resolve, `Added view: '${viewConfig.Title}'`, viewConfig.Title, view);
                            }).catch((error) => { Reject(reject, `Error while deleting all fields in the view with the title '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                        }).catch((error) => { Reject(reject, `Error while adding view with the title '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                    }
                    else if (result.length === 1) { Resolve(reject, `View with the title '${viewConfig.Title}' already exists`, viewConfig.Title, view); }
                    else { Reject(reject, `Found more than one view with the title '${viewConfig.Title}'`, viewConfig.Title, view); }
                }).catch((error) => { Reject(reject, `Error while requesting view with the title '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
            });
        });
    }

    private UpdateView(viewConfig: IView, parentPromise: Promise<List>): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            parentPromise.then((parentInstance => {
                Logger.write(`Updating View: '${viewConfig.Title}'`, Logger.LogLevel.Info);
                let view = parentInstance.views.getByTitle(viewConfig.Title);
                parentInstance.views.filter(`Title eq '${viewConfig.Title}'`).select("Id").get().then((result) => {
                    if (result.length === 1) {
                        let properties = this.CreateProperties(viewConfig);
                        view.update(properties).then((result) => {
                            view.fields.removeAll().then((result) => {
                                Resolve(resolve, `Updated view: '${viewConfig.Title}'`, viewConfig.Title, view);
                            }).catch((error) => { Reject(reject, `Error while deleting all fields in the view with the title '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                        }).catch((error) => { Reject(reject, `Error while updating view with the title '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                    }
                    else if (result.length === 0) { Reject(reject, `View with title '${viewConfig.Title}' does not exists`, viewConfig.Title, view); }
                }).catch((error) => { Reject(reject, `Error while requesting view with the title '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
            }));
        });
    }

    private DeleteView(viewConfig: IView, parentPromise: Promise<List>): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            parentPromise.then((parentInstance) => {
                Logger.write(`Deleting view: '${viewConfig.Title}'`, Logger.LogLevel.Info);
                let view = parentInstance.views.getByTitle(viewConfig.Title);
                parentInstance.views.filter(`Title eq '${viewConfig.Title}'`).select("Id").get().then((result) => {
                    if (result.length === 1) {
                        view.delete().then(() => {
                            Resolve(resolve, `Deleted view: '${viewConfig.Title}'`, viewConfig.Title, view);
                        }).catch((error) => { Reject(reject, `Error while deleting view with the title '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                    }
                    else if (result.length === 0) { Reject(reject, `View with the title '${viewConfig.Title}' does not exist`, viewConfig.Title, view); }
                }).catch((error) => { Reject(reject, `Error while requesting view with the title '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
            });
        });
    }

    private CreateProperties(viewConfig: IView) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(viewConfig);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (viewConfig.ControlOption) {
            case "Update":
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