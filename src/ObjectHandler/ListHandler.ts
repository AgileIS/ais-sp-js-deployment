import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import { Logger } from "sp-pnp-js/lib/utils/logging";
import { Web } from "sp-pnp-js/lib/sharepoint/rest/webs";
import { List, Lists } from "sp-pnp-js/lib/sharepoint/rest/lists";
import { Queryable } from "sp-pnp-js/lib/sharepoint/rest/queryable";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IList } from "../interface/Types/IList";
import { ListTemplates } from "../Constants/ListTemplates";
import { Resolve, Reject } from "../Util/Util";

export class ListHandler implements ISPObjectHandler {
    execute(listConfig: IList, parentPromise: Promise<Web>): Promise<List> {
        switch (listConfig.ControlOption) {
            case "Update":
                return this.UpdateList(listConfig, parentPromise);
            case "Delete":
                return this.DeleteList(listConfig, parentPromise);
            default:
                return this.AddList(listConfig, parentPromise);
        }
    }

    private AddList(listConfig: IList, parentPromise: Promise<Web>): Promise<List> {
        //todo: get list status in execute
        return new Promise<List>((resolve, reject) => {
            parentPromise.then(parentInstance => {
                Logger.write(`Adding list: '${listConfig.Title}'`, Logger.LogLevel.Info);
                if (listConfig.TemplateType) {
                    parentInstance.lists.filter(`RootFolder/Name eq '${listConfig.InternalName}'`).select("Id").get().then((result) => {
                        let list = undefined;
                        if (result.length === 0) {
                            let properties = this.CreateProperties(listConfig);
                            parentInstance.lists.add(listConfig.InternalName, listConfig.Description, listConfig.TemplateType, listConfig.EnableContentTypes, properties).then((result) => {
                                let list = parentInstance.lists.getById(result.data.Id);
                                //todo: use UpdateList
                                list.update({ Title: listConfig.Title }).then((result) => {
                                    Resolve(resolve, `Added list: '${listConfig.Title}'`, listConfig.Title, list);
                                }).catch((error) => { Reject(reject, `Error while updating list title with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
                            }).catch((error) => { Reject(reject, `Error while adding list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
                        } else {
                            list = parentInstance.lists.getById(result[0].Id);
                            Resolve(resolve, `List with Internal Name '${listConfig.InternalName}' already exists`, listConfig.InternalName, list);
                        }
                    }).catch((error) => { Reject(reject, `Error while requesting list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
                }
                else { Reject(reject, `List Template Type could not be resolved for the list with the internal name ${listConfig.InternalName}:`, listConfig.Title); }
            });
        });
    }

    private UpdateList(listConfig: IList, parentPromise: Promise<Web>): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            parentPromise.then(parentInstance => {
                Logger.write(`Updating View: '${listConfig.Title}'`, Logger.LogLevel.Info);
                parentInstance.lists.filter(`RootFolder/Name eq '${listConfig.InternalName}'`).select("Id").get().then((result) => {
                    if (result.length === 1) {
                        let list = parentInstance.lists.getById(result[0].Id);
                        let properties = this.CreateProperties(listConfig);
                        list.update(properties).then(() => {
                            Resolve(resolve, `Updated list: '${listConfig.Title}'`, listConfig.Title, list);
                        }).catch((error) => { Reject(reject, `Error while updating list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title, list); });
                    }
                    else { Reject(reject, `List with the internal name '${listConfig.InternalName}' does not exists`, listConfig.Title); }
                }).catch((error) => { Reject(reject, `Error while requesting list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
            });
        });
    }

    private DeleteList(listConfig: IList, parentPromise: Promise<Web>): Promise<List> {
        return new Promise<List>((resolve, reject) => {
            parentPromise.then(parentInstance => {
                Logger.write(`Deleting list: '${listConfig.Title}'`, Logger.LogLevel.Info);
                parentInstance.lists.filter(`RootFolder/Name eq '${listConfig.InternalName}'`).select("Id").get().then((result) => {
                    if (result.length === 1) {
                        let list = parentInstance.lists.getById(result[0].Id);
                        list.delete().then(() => {
                            Resolve(resolve, `Deleted List: '${listConfig.InternalName}'`, listConfig.Title, list);
                        }).catch((error) => { Reject(reject, `Error while deleting list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title, list); });
                    }
                    else { Reject(reject, `List with the internal name '${listConfig.InternalName}' does not exists`, listConfig.Title); }
                }).catch((error) => { Reject(reject, `Error while requesting list with the internal name '${listConfig.InternalName}': ` + error, listConfig.Title); });
            });
        });
    }

    private CreateProperties(listConfig: IList) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(listConfig);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (listConfig.ControlOption) {
            case "Update":
                delete parsedObject.InternalName;
                delete parsedObject.ControlOption;
                delete parsedObject.Field;
                delete parsedObject.EnableContentTypes;
                delete parsedObject.View;
                delete parsedObject.TemplateType;
                break;
            default:
                delete parsedObject.InternalName;
                delete parsedObject.EnableContentTypes;
                delete parsedObject.ControlOption;
                delete parsedObject.Field;
                delete parsedObject.View;
                delete parsedObject.Title;
                delete parsedObject.Description;
                delete parsedObject.TemplateType;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }

}