import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IList, } from "../interface/Types/IList";
import {ListTemplates} from "../Constants/ListTemplates";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import {Web} from "sp-pnp-js/lib/sharepoint/rest/webs";
import {List, Lists} from "sp-pnp-js/lib/sharepoint/rest/lists";
import {Queryable} from "sp-pnp-js/lib/sharepoint/rest/queryable";
import {RejectAndLog} from "../Util/Util";

export class ListHandler implements ISPObjectHandler {

    execute(config: IList, parent: Promise<Web>): Promise<List> {
        switch (config.ControlOption) {
            case "Update":
                return this.UpdateList(config, parent);
            case "Delete":
                return this.DeleteList(config, parent);
            default:
                return this.AddList(config, parent);
        }
    }

    private AddList(config: IList, parent: Promise<Web>): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            parent.then(parentInstance => {
                Logger.write("config " + JSON.stringify(config), 0);
                if (config.TemplateType) {
                    parentInstance.lists.filter(`RootFolder/Name eq '${config.InternalName}'`).select("Id").get().then((result) => {
                        if (result.length === 0) {
                            let propertyHash = this.CreateProperties(config);
                            parentInstance.lists.add(config.InternalName, config.Description, config.TemplateType, config.EnableContentTypes, propertyHash).then((result) => {
                                let listId = result.data.Id;
                                result.list.update({ Title: config.Title }).then((result) => {
                                    let list = parentInstance.lists.getById(listId);
                                    resolve(list);
                                    Logger.write(`List ${config.Title} created`, 0);
                                }, (error) => {
                                    RejectAndLog(error, config.Title, reject);
                                });
                            }, (error) => {
                                reject(error);
                            });
                        } else {
                            let list = parentInstance.lists.getById(result[0].Id);
                            resolve(list);
                            Logger.write(`List with Internal Name '${config.InternalName}' already exists`, 0);
                        }
                    }, (error) => {
                        RejectAndLog(error, config.Title, reject);
                    });
                }
                else {
                    let error = `List Template Type could not be resolved for List: ${config.InternalName}`;
                    RejectAndLog(error, config.Title, reject);
                }
            });
        });
    }

    private UpdateList(config: IList, parent: Promise<Web>): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            parent.then(parentInstance => {
                Logger.write("config " + JSON.stringify(config), 0);
                let list = parentInstance.lists.getByTitle(config.Title);
                parentInstance.lists.filter(`RootFolder/Name eq '${config.InternalName}'`).select("Id").get().then((result) => {
                    if (result.length === 1) {
                        let listId = result[0].Id;
                        let properties = this.CreateProperties(config);
                        list.update(properties).then(() => {
                            resolve(list);
                            Logger.write(`List with Internal Name '${config.InternalName}' updated`, 1);
                        }).catch((error) => {
                            RejectAndLog(error, config.Title, reject);
                        });
                    }
                    else {
                        let error = `List with Internal Name '${config.InternalName}' does not exist`;
                        RejectAndLog(error, config.Title, reject);
                    }
                }).catch((error) => {
                    RejectAndLog(error, config.Title, reject);
                });
            })

        });
    }

    private DeleteList(config: IList, parent: Promise<Web>): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            parent.then(parentInstance => {
                Logger.write("config " + JSON.stringify(config), 0);
                let list = parentInstance.lists.getByTitle(config.Title);
                parentInstance.lists.filter(`RootFolder/Name eq '${config.InternalName}'`).select("Id").get().then((result) => {
                    if (result.length === 1) {
                        let listId = result[0].Id;
                        list.delete().then(() => {
                            let configForDelete = this.CreateProperties(config);
                            resolve(configForDelete);
                            Logger.write(`List with Internal Name '${config.InternalName}' deleted`, 1);
                        }).catch((error) => {
                            RejectAndLog(error, config.Title, reject);
                        });
                    }
                    else {
                        let error = `List with Internal Name '${config.InternalName}' does not exist`;
                        RejectAndLog(error, config.Title, reject);
                    }
                }).catch((error) => {
                    RejectAndLog(error, config.Title, reject);
                });
            })

        });
    }


    private CreateProperties(pElement: IList) {
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

}