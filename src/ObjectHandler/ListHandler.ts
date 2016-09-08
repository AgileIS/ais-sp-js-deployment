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
        /* return new Promise<List>((resolve, reject) => {
             parent.then((parentInstance) => {
                 Logger.write("enter List execute", 0);
                 let list = parentInstance.lists.getByTitle(config.Title);
                 list.get().then(result => {
                     resolve(list);
                 });
             });
 
         });*/

        switch (config.ControlOption) {
            case "Update":
                return this.UpdateList(config, parent);
            case "Delete":
                return this.DeleteList(config, parent);
            default:
                return this.AddList(config, parent);
        }

    }
    /*    execute(config: IList, url: string, parentConfig: ISite) {
            switch (config.ControlOption) {
                case "Update":
                    return this.UpdateList(config, url);
                case "Delete":
                    return this.DeleteList(config, url);
    
                default:
                    return this.AddList(config, url);
            }
        }*/

    private AddList(config: IList, parent: Promise<Web>): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            parent.then(parentInstance => {
                Logger.write("config " + JSON.stringify(config), 0);
                let list = parentInstance.lists.getByTitle(config.Title);
                if (list !== undefined) {
                    if (config.TemplateType) {
                        parentInstance.lists.filter(`RootFolder/Name eq '${config.InternalName}'`).select("Id").get().then((result) => {
                            if (result.length === 0) {
                                let propertyHash = this.CreateProperties(config);
                                parentInstance.lists.add(config.InternalName, config.Description, config.TemplateType, config.EnableContentTypes, propertyHash).then((result) => {
                                    result.list.update({ Title: config.Title }).then((result) => {
                                        resolve(list);
                                        Logger.write(`List ${config.Title} created`, 0);
                                    }, (error) => {
                                        reject(error);
                                    });
                                }, (error) => {
                                    reject(error);
                                });
                            } else {
                                resolve(list);
                                Logger.write(`List with Internal Name '${config.InternalName}' already exists`, 0);
                            }
                        }, (error) => {
                            reject(error);
                        });
                    }
                    else {
                        let error = `List Template Type could not be resolved for List: ${config.InternalName}`;
                        reject(error);
                    }
                }

            })

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
                            reject(error);
                        });
                    }
                    else {
                        let error = `List with Internal Name '${config.InternalName}' does not exist`;
                        reject(error);
                    }
                }).catch((error) => {
                    reject(error);
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
                            reject(error);
                        });
                    }
                    else {
                        let error = `List with Internal Name '${config.InternalName}' does not exist`;
                        reject(error);
                    }
                }).catch((error) => {
                    reject(error);
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