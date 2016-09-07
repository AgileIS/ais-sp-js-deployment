import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IField} from "../interface/Types/IField";
import {IList} from "../interface/Types/IList";
import {ISite} from "../interface/Types/ISite";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import {Web} from "sp-pnp-js/lib/sharepoint/rest/webs";
import {RejectAndLog} from "../lib/Util/Util";

export class FieldHandler implements ISPObjectHandler {
    public execute(config: IField, url: string, parentConfig: ISite | IList) {
        switch (config.ControlOption) {
            case "Update":
                return this.UpdateField(config, url);
            case "Delete":
                return this.DeleteField(config, url);
            default:
                return this.AddField(config, url, parentConfig);
        }
    }

    private AddField(config: IField, url: string, parentConfig: ISite | IList) {
        let spWeb = new Web(url);
        let element = config;
        return new Promise<IField>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config));
            let parentObject = parentConfig as Object;
            if (parentObject.hasOwnProperty("List")) {
                this.checkSiteField(url, element).then(() => {
                    resolve(element);
                }).catch((error) => {
                    RejectAndLog(error, element.InternalName, reject);
                });
            } else {
                let parentElement = parentObject as IList;
                this.checkListField(url, element, parentElement).then((result) => {
                    resolve(element);
                }).catch((error) => {
                    RejectAndLog(error, element.InternalName, reject);
                });
            };
        });
    };
    private UpdateField(config: IField, url: string) {
        let spWeb = new Web(url);
        let element = config;
        return new Promise<IField>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config), 0);
            spWeb.fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let fieldId = result[0].Id;
                    let properties = this.CreateProperties(element);
                    spWeb.fields.getById(fieldId).update(properties).then(() => {
                        resolve(element);
                        Logger.write(`Field with Internal Name '${element.InternalName}' updated`, 1);
                    }).catch((error) => {
                        RejectAndLog(error, element.InternalName, reject);
                    });
                }
                else {
                    let error = `Field with Internal Name '${element.InternalName}' does not exist`;
                    RejectAndLog(error, element.InternalName, reject);
                }
            });
        });
    }

    private DeleteField(config: IField, url: string) {
        let spWeb = new Web(url);
        let element = config;
        return new Promise<IField>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config), 0);
            spWeb.fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let fieldId = result[0].Id;
                    spWeb.fields.getById(fieldId).delete().then(() => {
                        Logger.write(`Field with Internal Name '${element.InternalName}' deleted`, 1);
                        resolve(element);
                    }).catch((error) => {
                        RejectAndLog(error, element.InternalName, reject);
                    });
                }
                else {
                    let error = `Field with Internal Name '${element.InternalName}' does not exist`;
                    RejectAndLog(error, element.InternalName, reject);
                }
            });
        });
    }

    private checkListField(url: string, pConfig: IField, pParentElement: IList): Promise<any> {
        return new Promise((resolve, reject) => {
            let spWeb = new Web(url);
            let element = pConfig;
            let listName = pParentElement.InternalName;
            spWeb.lists.filter(`RootFolder/Name eq '${listName}'`).select("Id").get().then((data) => {
                if (data.length === 1) {
                    let listId = data[0].Id;
                    spWeb.lists.getById(listId).fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                        if (result.length === 0) {
                            if (element.FieldTypeKind) {
                                if (element.FieldTypeKind === 7) {
                                    // Lookup nicht vorhanden umsetzten als create FieldAsXml
                                } else if (element.FieldTypeKind === 17) {
                                    let propertyHash = this.CreateProperties(element);
                                    spWeb.lists.getById(listId).fields.addCalculated(element.InternalName, element.Formula, Types.DateTimeFieldFormatType[element.DateFormat], Types.FieldTypes[element.OutputType], propertyHash).then((result) => {
                                        result.field.update({ Title: element.Title, Description: element.Description }).then(() => {
                                            resolve();
                                            Logger.write("Calculated Field with Internal Name '" + element.InternalName + "' created", 1);
                                        }).catch((error) => {
                                            reject(error);
                                        });
                                    }).catch((error) => {
                                        reject(error);
                                    });
                                } else {
                                    let propertyHash = this.CreateProperties(element);
                                    spWeb.lists.getById(listId).fields.add(element.InternalName, "SP.Field", propertyHash).then((result) => {
                                        result.field.update({ Title: element.Title }).then(() => {
                                            Logger.write(`Field with Internal Name '${element.InternalName}' created`);
                                            resolve();
                                        }).catch((error) => {
                                            reject(error);
                                        });
                                    }).catch((error) => {
                                        reject(error);
                                    });
                                }
                            } else {
                                let error = `FieldTypKind for '${element.InternalName}' could not be resolved`;
                                reject(error);
                            }
                        } else {
                            resolve(listId);
                            Logger.write(`Field with Internal Name '${element.InternalName}' already exists`);
                        }
                    }).catch((error) => {
                        reject(error);
                    });
                } else {
                    let error = `List with Title '${listName}' for Field '${element.InternalName}' does not exist`;
                    reject(error);
                }
            }).catch((error) => {
                reject(error);
            });
        });
    }

    private checkSiteField(url: string, pConfig: IField): Promise<any> {
        return new Promise((resolve, reject) => {
            let spWeb = new Web(url);
            let element = pConfig;
            spWeb.fields.filter("InternalName eq '" + element.InternalName + "'").select("Id").get().then((data) => {
                if (data.length === 0) {
                    if (element.FieldTypeKind) {
                        if (element.FieldTypeKind === 7) {  // 7 = Lookup
                            resolve(element);
                        }
                        else if (element.FieldTypeKind === 17) { // 17 = Calculated
                            let propertyHash = this.CreateProperties(element);
                            spWeb.fields.addCalculated(element.InternalName, element.Formula, Types.DateTimeFieldFormatType[element.DateFormat], Types.FieldTypes[element.OutputType], propertyHash).then((result) => {
                                result.field.update({ Title: element.Title, Description: element.Description }).then(() => {
                                    resolve();
                                    Logger.write("Calculated Field with Internal Name '" + element.InternalName + "' created", 1);
                                }).catch((error) => {
                                    reject(error);
                                });
                            }).catch((error) => {
                                reject(error);
                            });
                        }
                        else {
                            let propertyHash = this.CreateProperties(element);
                            spWeb.fields.add(element.InternalName, "SP.Field", propertyHash).then((result) => {
                                result.field.update({ Title: element.Title }).then(() => {
                                    resolve();
                                    Logger.write("Field with Internal Name'" + element.InternalName + "' created", 1);
                                }).catch((error) => {
                                    reject(error);
                                });
                            }).catch((error) => {
                                reject(error);
                            });
                        }
                    }
                    else {
                        let error = "FieldTypKind could not be resolved";
                        reject(error);
                    }
                }
                else {
                    let error = "Field with Internal Name '" + element.InternalName + "' already exists";
                    reject(error);
                }
            }).catch((error) => {
                reject(error);
            });
        });
    }



    private CreateProperties(pElement: IField) {
        let element = pElement;
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(element);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (element.ControlOption) {
            case "":
                delete parsedObject.ControlOption;
                delete parsedObject.Title;
                delete parsedObject.Description;
                delete parsedObject.DateFormat;
                delete parsedObject.Formula;
                delete parsedObject.OutputType;
                break;
            case "Update":
                delete parsedObject.ControlOption;
                delete parsedObject.InternalName;
                delete parsedObject.FieldTypeKind;
                delete parsedObject.DateFormat;
                delete parsedObject.OutputType;
                delete parsedObject.Formula;
                break;
            default:
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}

