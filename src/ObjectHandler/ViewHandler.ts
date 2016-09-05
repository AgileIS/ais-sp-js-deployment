import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {IView} from "../interface/Types/IView";
import {IList} from "../interface/Types/IList";
import {RejectAndLog} from "../lib/Util/Util";


export class ViewHandler implements ISPObjectHandler {
    execute(config: IView, url: string, parentConfig: IList) {
        switch (config.ControlOption) {
            case "":
                return AddView(config, url, parentConfig);
            case "Update":
                return UpdateView(config, url, parentConfig);
            case "Delete":
                return DeleteView(config, url, parentConfig);

            default:
                return AddView(config, url, parentConfig);
        }
    };
}

function AddView(config: IView, url: string, parentConfig: IList) {
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
                            let properties = CreateProperties(element);
                            spWeb.lists.getById(listId).views.add(element.Title, element.PersonalView, properties).then(
                                (result) => {
                                    result.view.fields.removeAll().then(() => {
                                        let configForAddView = AddListNameProperty(element, listId);
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

function UpdateView(config: IView, url: string, parentConfig: IList) {
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
                        let properties = CreateProperties(element);
                        spWeb.lists.getById(listId).views.getById(viewId).update(properties).then(() => {
                            spWeb.lists.getById(listId).views.getById(viewId).fields.removeAll().then(() => {
                                let configForAddView = AddListNameProperty(element, listId);
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


function DeleteView(config: IView, url: string, parentConfig: IList) {
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
                            let configForDelete = CreateProperties(element);
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




function CreateProperties(pElement: IView) {
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

function AddListNameProperty(pElement: IView, pParentListId: any): IView {
    let element = pElement;
    let stringifiedObject: string;
    stringifiedObject = JSON.stringify(element);
    let parsedObject = JSON.parse(stringifiedObject);
    parsedObject.ParentListId = pParentListId;
    stringifiedObject = JSON.stringify(parsedObject);
    return JSON.parse(stringifiedObject);
}