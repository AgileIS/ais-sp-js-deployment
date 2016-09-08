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

    execute(config: IList, parent: Promise<Web>):Promise<List> {
        return new Promise<List>((resolve, reject) => {
            parent.then((parentInstance) => {
                Logger.write("enter List execute", 0);
                // das hier geht noch nicht INstance eventuell überdenken!
                let list = parentInstance.lists.getByTitle(config.Title);  
               list.get().then(result => {
                    resolve(list);
                });
            });

        });

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

    private AddList(config: IList, url: string) {
        let element = config;
        let spWeb = new Web(url);
        return new Promise<IList>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config), 0);
            if (element.TemplateType) {
                spWeb.lists.filter(`RootFolder/Name eq '${element.InternalName}'`).select("Id").get().then((result) => {
                    if (result.length === 0) {
                        let propertyHash = this.CreateProperties(element);
                        spWeb.lists.add(element.InternalName, element.Description, element.TemplateType, element.EnableContentTypes, propertyHash).then((result) => {
                            result.list.update({ Title: element.Title }).then((result) => {
                                resolve(element);
                                Logger.write(`List ${element.Title} created`, 0);
                            }, (error) => {
                                RejectAndLog(error, element.InternalName, reject);
                            });
                        }, (error) => {
                            RejectAndLog(error, element.InternalName, reject);
                        });
                    } else {
                        resolve(element);
                        Logger.write(`List with Internal Name '${element.InternalName}' already exists`, 0);
                    }
                }, (error) => {
                    RejectAndLog(error, element.InternalName, reject);
                });
            }
            else {
                let error = `List Template Type could not be resolved for List: ${element.InternalName}`;
                RejectAndLog(error, element.InternalName, reject);
            }
        });
    }

    private UpdateList(config: IList, url: string) {
        let element = config;
        let spWeb = new Web(url);
        return new Promise<IList>((resolve, reject) => {
            spWeb.lists.filter(`RootFolder/Name eq '${element.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let listId = result[0].Id;
                    let properties = this.CreateProperties(element);
                    spWeb.lists.getById(listId).update(properties).then(() => {
                        resolve(element);
                        Logger.write(`List with Internal Name '${element.InternalName}' updated`, 1);
                    }).catch((error) => {
                        RejectAndLog(error, element.InternalName, reject);
                    });
                }
                else {
                    let error = `List with Internal Name '${element.InternalName}' does not exist`;
                    RejectAndLog(error, element.InternalName, reject);
                }
            }).catch((error) => {
                RejectAndLog(error, element.InternalName, reject);
            });
        });
    }

    private DeleteList(config: IList, url: string) {
        let element = config;
        let spWeb = new Web(url);
        return new Promise<IList>((resolve, reject) => {
            spWeb.lists.filter(`RootFolder/Name eq '${element.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let listId = result[0].Id;
                    spWeb.lists.getById(listId).delete().then(() => {
                        let configForDelete = this.CreateProperties(config);
                        resolve(configForDelete);
                        Logger.write(`List with Internal Name '${element.InternalName}' deleted`, 1);
                    }).catch((error) => {
                        RejectAndLog(error, element.InternalName, reject);
                    });
                }
                else {
                    let error = `List with Internal Name '${element.InternalName}' does not exist`;
                    RejectAndLog(error, element.InternalName, reject);
                }
            }).catch((error) => {
                RejectAndLog(error, element.InternalName, reject);
            });
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