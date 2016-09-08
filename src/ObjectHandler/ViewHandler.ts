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
        return new Promise<View>((resolve, reject) => {
            parentPromise.then((parentInstance) => {
                Logger.write(`Adding View: '${viewConfig.Title}'`, Logger.LogLevel.Info);
                let view = parentInstance.views.getByTitle(viewConfig.Title);
                view.get().then((result) => {
                    if (result.length === 0) {
                        let properties = this.CreateProperties(viewConfig);
                        parentInstance.views.add(viewConfig.Title, viewConfig.PersonalView, properties).then((result) => {
                            result.view.fields.removeAll().then(() => {
                                Resolve(resolve, `View '${viewConfig.Title}' added`, viewConfig.Title, view);
                            }).catch((error) => { Reject(reject, `Error while deleting all view fields '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                        }).catch((error) => { Reject(reject, `Error while adding view '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                    }
                    else if (result.length === 1) { Reject(reject, `View '${viewConfig.Title}' already exists`, viewConfig.Title, view); }
                    else { Reject(reject, `Found more than one view with the title '${viewConfig.Title}'`, viewConfig.Title, view); }
                }).catch((error) => { Reject(reject, `Error while requesting view '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
            }).catch((error) => { Reject(reject, error, viewConfig.Title); });
        });
    }

    private UpdateView(viewConfig: IView, parentPromise: Promise<List>): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            parentPromise.then((parentInstance => {
                Logger.write(`Updating View: '${viewConfig.Title}'`, Logger.LogLevel.Info);
                let view = parentInstance.views.getByTitle(viewConfig.Title);
                view.get().then((result) => {
                    if (result.length === 1) {
                        let properties = this.CreateProperties(viewConfig);
                        view.update(properties).then((result) => {
                            view.fields.removeAll().then((result) => {
                                Resolve(resolve, `View '${viewConfig.Title}' updated`, viewConfig.Title, view);
                            }).catch((error) => { Reject(reject, `Error while deleting all view fields '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                        }).catch((error) => { Reject(reject, `Error while updating view '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                    }
                    else if (result.length === 0) { Reject(reject, `View '${viewConfig.Title}' does not exists`, viewConfig.Title, view); }
                }).catch((error) => { Reject(reject, `Error while requesting view '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
            })).catch((error) => { Reject(reject, error, viewConfig.Title); });
        });
    }

    private DeleteView(viewConfig: IView, parentPromise: Promise<List>): Promise<View> {
        return new Promise<View>((resolve, reject) => {
            parentPromise.then((parentInstance) => {
                Logger.write(`Deleting View '${viewConfig.Title}'`, Logger.LogLevel.Info);
                let view = parentInstance.views.getByTitle(viewConfig.Title);
                view.get().then((result) => {
                    if (result.length === 1) {
                        view.delete().then(() => {
                            Resolve(resolve, `View '${viewConfig.Title}' removed`, viewConfig.Title, view);
                        }).catch((error) => { Reject(reject, `Error while deleting view '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
                    }
                    else if (result.length === 0) { Reject(reject, `View '${viewConfig.Title}' does not exist`, viewConfig.Title, view); }
                }).catch((error) => { Reject(reject, `Error while requesting view '${viewConfig.Title}': ` + error, viewConfig.Title, view); });
            }).catch((error) => { Reject(reject, error, viewConfig.Title); });
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
            case "Delete":
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