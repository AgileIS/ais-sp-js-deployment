import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {IView} from "../interface/Types/IView";
import {IList} from "../interface/Types/IList";
import {RejectAndLog} from "../lib/Util/Util";


export class ViewHandler implements ISPObjectHandler {
    public execute(config: IView, url: string, parentConfig: IList): Promise<IView> {
        switch (config.ControlOption) {
            case "Update":
                return this.UpdateView(config, url, parentConfig);
            case "Delete":
                return this.DeleteView(config, url, parentConfig);
            default:
                return this.AddView(config, url, parentConfig);
        }
    };

    private AddView(config: IView, url: string, parentConfig: IList): Promise<IView> {
        let spWeb = new web.Web(url);
        let element = config;
        let listName = parentConfig.InternalName;
        return new Promise<IView>((resolve, reject) => {
            spWeb.lists.filter(`RootFolder/Name eq '${listName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let listId = result[0].Id;
                    spWeb.lists.getById(listId).views.filter(`Title eq '${element.Title}'`).select("Id").get().then(
                        (result) => {
                            if (result.length === 1) {
                                resolve(element);
                                Logger.write(`View with Title '${element.Title}' already exists`, 1);
                            }
                            else if (result.length === 0) {
                                let properties = this.CreateProperties(element);
                                spWeb.lists.getById(listId).views.add(element.Title, element.PersonalView, properties).then(
                                    (result) => {
                                        result.view.fields.removeAll().then(() => {
                                            let configForAddView = this.AddListNameProperty(element, listId);
                                            resolve(configForAddView);
                                        }).catch((error) => {
                                            RejectAndLog(error, element.Title, reject);
                                        });
                                    }).catch((error) => {
                                        RejectAndLog(error, element.Title, reject);
                                    });
                            }
                            else {
                                let error = `More than one Views wit Title '${element.Title}' found`;
                                RejectAndLog(error, element.Title, reject);
                            }
                        }).catch((error) => {
                            RejectAndLog(error, element.Title, reject);
                        });
                }
                else {
                    let error = `List with Internal Name '${listName}' does not exist`;
                    RejectAndLog(error, element.Title, reject);
                }
            });
        });
    }

    private UpdateView(config: IView, url: string, parentConfig: IList): Promise<IView> {
        let spWeb = new web.Web(url);
        let element = config;
        let listName = parentConfig.InternalName;
        return new Promise<IView>((resolve, reject) => {
            spWeb.lists.filter(`RootFolder/Name eq '${listName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let listId = result[0].Id;
                    spWeb.lists.getById(listId).views.filter(`Title eq '${element.Title}'`).select("Id").get().then((result) => {
                        if (result.length === 1) {
                            let viewId = result[0].Id;
                            let properties = this.CreateProperties(element);
                            spWeb.lists.getById(listId).views.getById(viewId).update(properties).then(() => {
                                spWeb.lists.getById(listId).views.getById(viewId).fields.removeAll().then(() => {
                                    let configForAddView = this.AddListNameProperty(element, listId);
                                    resolve(configForAddView);
                                }).catch((error) => {
                                    RejectAndLog(error, element.Title, reject);
                                });
                            });
                        }
                        else if (result.length === 0) {
                            let error = `View with Title '${element.Title}' does not exist`;
                            RejectAndLog(error, element.Title, reject);
                        }
                    });
                }
                else {
                    let error = `List with Internal Name '${listName}' does not exist`;
                    RejectAndLog(error, element.Title, reject);
                }
            }).catch((error) => {
                RejectAndLog(error, element.Title, reject);
            });
        });
    }


    private DeleteView(config: IView, url: string, parentConfig: IList): Promise<IView> {
        let spWeb = new web.Web(url);
        let element = config;
        let listName = parentConfig.InternalName;
        return new Promise<IView>((resolve, reject) => {
            spWeb.lists.filter(`RootFolder/Name eq '${listName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let listId = result[0].Id;
                    spWeb.lists.getById(listId).views.filter(`Title eq '${element.Title}'`).select("Id").get().then((result) => {
                        if (result.length === 1) {
                            let viewId = result[0].Id;
                            spWeb.lists.getById(listId).views.getById(viewId).delete().then(() => {
                                let configForDelete = this.CreateProperties(element);
                                resolve(configForDelete);
                                Logger.write(`View with Title '${element.Title}' removed`, 1);
                            });
                        }
                        else if (result.length === 0) {
                            let error = `View with Title '${element.Title}' does not exist`;
                            RejectAndLog(error, element.Title, reject);
                        }
                    });
                }
                else {
                    let error = `List with Internal Name '${listName}' does not exist`;
                    RejectAndLog(error, element.Title, reject);

                }
            });
        });
    }



    private CreateProperties(pElement: IView) {
        let element = pElement;
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(element);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (element.ControlOption) {
            case "":
                delete parsedObject.ControlOption;
                delete parsedObject.Title;
                delete parsedObject.PersonalView;
                delete parsedObject.ViewField;
                break;
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
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }

    private AddListNameProperty(pElement: IView, pParentListId: any): IView {
        let element = pElement;
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(element);
        let parsedObject = JSON.parse(stringifiedObject);
        parsedObject.ParentListId = pParentListId;
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }

}