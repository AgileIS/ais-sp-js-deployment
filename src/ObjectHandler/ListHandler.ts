import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IList} from "../interface/Types/IList";
import {ISite} from "../interface/Types/ISite";
import {ListTemplates} from "../lib/ListTemplates";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";

export class ListHandler implements ISPObjectHandler {
    execute(config: IList, url: string, parentConfig: ISite) {
        switch (config.ControlOption) {
            case "":
                return AddList(config, url);
            case "Update":
                return UpdateList(config, url);
            case "Delete":
                return DeleteList(config, url);

            default:
                return Promise.reject("Control Option on Element not found");
        }
    }
}


function AddList(config: IList, url: string) {
    let element = config;
    let spWeb = new web.Web(url);
    return new Promise<IList>((resolve, reject) => {
        Logger.write("config " + JSON.stringify(config), 0);
        if (element.TemplateType) {
            spWeb.lists.filter(`RootFolder/Name eq '${element.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 0) {
                    let propertyHash = CreateProperties(element);
                    spWeb.lists.add(element.InternalName, element.Description, element.TemplateType, element.EnableContentTypes, propertyHash).then((result) => {
                        result.list.update({ Title: element.Title }).then((result) => {
                            resolve(config);
                            Logger.write(`List ${element.Title} created`, 0);
                        }, (error) => {
                            reject(error + " - " + element.InternalName);
                        });
                    }, (error) => {
                        reject(error + " - " + element.InternalName);
                    });
                } else {
                    resolve(config);
                    Logger.write(`List with Internal Name '${element.InternalName}' already exists`, 0);
                }
            }, (error) => {
                reject(error + " - " + element.InternalName);
            });
        }
        else {
            let error = `List Template Type could not be resolved for List: ${element.InternalName}`;
            reject(error);
        }
    });
}

function UpdateList(config: IList, url: string) {
    let element = config;
    let spWeb = new web.Web(url);
    return new Promise<IList>((resolve, reject) => {
        spWeb.lists.filter(`RootFolder/Name eq '${element.InternalName}'`).select("Id").get().then((result) => {
            if (result.length === 1) {
                let listId = result[0].Id;
                let properties = CreateProperties(element);
                spWeb.lists.getById(listId).update(properties).then(() => {
                    resolve(config);
                    Logger.write(`List with Internal Name '${element.InternalName}' updated`, 1);
                }).catch((error) => {
                    reject(error + " - " + element.InternalName);
                });
            }
            else {
                let error = `List with Internal Name '${element.InternalName}' does not exist`;
                reject(error);
            }
        }).catch((error) => {
            reject(error + " - " + element.InternalName);
        });
    });
}

function DeleteList(config: IList, url: string) {
    let element = config;
    let spWeb = new web.Web(url);
    return new Promise<IList>((resolve, reject) => {
        spWeb.lists.filter(`RootFolder/Name eq '${element.InternalName}'`).select("Id").get().then((result) => {
            if (result.length === 1) {
                let listId = result[0].Id;
                spWeb.lists.getById(listId).delete().then(() => {
                    let configForDelete = CreateProperties(config);
                    resolve(configForDelete);
                    Logger.write(`List with Internal Name '${element.InternalName}' deleted`, 1);
                }).catch((error) => {
                    reject(error + " - " + element.InternalName);
                });
            }
            else {
                let error = `List with Internal Name '${element.InternalName}' does not exist`;
                reject(error);
            }
        }).catch((error) => {
            reject(error + " - " + element.InternalName);
        });
    });
}


function CreateProperties(pElement: IList) {
    let element = pElement;
    let stringifiedObject: string;
    stringifiedObject = JSON.stringify(element);
    let parsedObject = JSON.parse(stringifiedObject);
    switch (element.ControlOption) {
        case "":
            delete parsedObject.InternalName;
            delete parsedObject.EnableContentTypes;
            delete parsedObject.ControlOption;
            delete parsedObject.Field;
            delete parsedObject.View;
            delete parsedObject.Title;
            delete parsedObject.Description;
            delete parsedObject.TemplateType;
            break;
        case "Update":
            delete parsedObject.InternalName;
            delete parsedObject.ControlOption;
            delete parsedObject.Field;
            delete parsedObject.EnableContentTypes;
            delete parsedObject.View;
            delete parsedObject.TemplateType;
            break;
        case "Delete":
            delete parsedObject.Field;
            delete parsedObject.View;
            delete parsedObject.ContentType;
            break;
        default:
            break;
    }
    stringifiedObject = JSON.stringify(parsedObject);
    return JSON.parse(stringifiedObject);
}