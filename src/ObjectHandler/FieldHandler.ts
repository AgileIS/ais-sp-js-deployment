import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IField}  from "../interface/Types/IField";
import {IList} from "../interface/Types/IList";
import {ISite} from "../interface/Types/ISite";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import {Fields} from "sp-pnp-js/lib/sharepoint/rest/Fields";
import {Web} from "sp-pnp-js/lib/sharepoint/rest/webs";
import {RejectAndLog} from "../Util/Util";
import {FieldTypeKind} from "../Constants/FieldTypeKind";

export class FieldHandler {
    execute(config: IField, url: string, parent: Promise<any>): Promise<any> {
        return new Promise<IField>((resolve, reject) => {
            parent.then((parentProperties) => {
                Logger.write("enter Field execute", 0);
                resolve(config);
            });

        });

    }

    /*   public execute(config: IField, url: string, parentConfig: ISite | IList) {
           return new Promise<IField>((resolve, reject) => {
               Logger.write("config " + JSON.stringify(config));
               let parentObject = parentConfig as Object;
               if (parentObject.hasOwnProperty("List")) {
                   switch (config.ControlOption) {
                       case "Update":
                           this.updateFieldOnSite(url, config).then(() => {
                               resolve(config);
                           }).catch((error) => {
                               RejectAndLog(error, config.InternalName, reject);
                           });
                           break;
                       case "Delete":
                           this.deleteFieldOnSite(config, url).then(() => {
                               resolve(config);
                           }).catch((error) => {
                               RejectAndLog(error, config.InternalName, reject);
                           });
                           break;
                       default:
                           this.addFieldSite(url, config).then(() => {
                               resolve(config);
                           }).catch((error) => {
                               RejectAndLog(error, config.InternalName, reject);
                           });
                           break;
                   }
               } else {
                   let parentElement = parentObject as IList;
                   switch (config.ControlOption) {
                       case "Update":
                           this.updateFieldOnList(url, config, parentElement).then(() => {
                               resolve(config);
                           }).catch((error) => {
                               RejectAndLog(error, config.InternalName, reject);
                           });
                           break;
                       case "Delete":
                           this.deleteFieldOnList(config, url, parentElement).then(() => {
                               resolve(config);
                           }).catch((error) => {
                               RejectAndLog(error, config.InternalName, reject);
                           });
                           break;
                       default:
                           this.addFieldOnList(url, config, parentElement).then((result) => {
                               resolve(config);
                           }).catch((error) => {
                               RejectAndLog(error, config.InternalName, reject);
                           });
                           break;
                   }
   
               };
           });
       }
   */
    private updateFieldOnSite(url: string, pConfig: IField): Promise<any> {
        let spWeb = new Web(url);
        let element = pConfig;
        return new Promise<IField>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(element), 0);
            spWeb.fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let fieldId = result[0].Id;
                    let properties = this.CreateProperties(element);
                    spWeb.fields.getById(fieldId).update(properties).then(() => {
                        resolve();
                        Logger.write(`Field with Internal Name '${element.InternalName}' updated`, 1);
                    }).catch((error) => {
                        reject(error);
                    });
                }
                else {
                    let error = `Field with Internal Name '${element.InternalName}' does not exist`;
                    reject(error);
                }
            });
        });
    }

    private updateFieldOnList(url: string, pConfig: IField, pParentElement: IList): Promise<any> {
        let spWeb = new Web(url);
        let element = pConfig;
        let listName: string = pParentElement.InternalName;
        return new Promise((resolve, reject) => {
            spWeb.lists.filter(`RootFolder/Name eq '${listName}'`).get().then((data) => {
                if (data.length === 1) {
                    let listId = data[0].Id;
                    spWeb.lists.getById(listId).fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                        if (result.length === 1) {
                            let fieldId = result[0].Id;
                            let properties = this.CreateProperties(element);
                            if (properties) {
                                spWeb.lists.getById(listId).fields.getById(fieldId).update(properties).then(() => {
                                    resolve();
                                    Logger.write(`Field with Internal Name '${element.InternalName}' updated`);
                                }).catch((error) => {
                                    reject(error);
                                });
                            } else {
                                let error = `No changes on Field '${element.InternalName}' found`;
                                reject(error);
                            }
                        } else {
                            let error = `Field with Internal Name '${element.Title}' does not exist`;
                            reject(error);
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
    private deleteFieldOnSite(config: IField, url: string) {
        let spWeb = new Web(url);
        let element = config;
        return new Promise<IField>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config), 0);
            spWeb.fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let fieldId = result[0].Id;
                    spWeb.fields.getById(fieldId).delete().then(() => {
                        Logger.write(`Field with Internal Name '${element.InternalName}' deleted`, 1);
                        resolve();
                    }).catch((error) => {
                        reject(error);
                    });
                }
                else {
                    let error = `Field with Internal Name '${element.InternalName}' does not exist`;
                    reject(error);
                }
            });
        });
    }
    private deleteFieldOnList(config: IField, url: string, pParentElement: IList) {
        let spWeb = new Web(url);
        let element = config;
        let listName = pParentElement.InternalName;
        return new Promise((resolve, reject) => {
            spWeb.lists.filter(`RootFolder/Name eq '${listName}'`).get().then((data) => {
                if (data.length === 1) {
                    let listId = data[0].Id;
                    spWeb.lists.getById(listId).fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                        if (result.length === 1) {
                            let fieldId = result[0].Id;
                            spWeb.lists.getById(listId).fields.getById(fieldId).delete().then(() => {
                                resolve();
                                Logger.write(`Field with Internal Name '${element.InternalName}' deleted`);
                            }).catch((error) => {
                                reject(error);
                            });
                        } else {
                            let error = `Field with Internal Name '${element.Title}' does not exist`;
                            reject(error);
                        }
                    });
                } else {

                    let error = `List with Title '${listName}' for Field '${element.InternalName}' does not exist`;
                    reject(error);
                }
            });
        });
    }

    private addFieldOnList(url: string, pConfig: IField, pParentElement: IList): Promise<any> {
        return new Promise((resolve, reject) => {
            let spWeb = new Web(url);
            let element = pConfig;
            let listName = pParentElement.InternalName;
            spWeb.lists.filter(`RootFolder/Name eq '${listName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let listId = result[0].Id;
                    spWeb.lists.getById(listId).fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                        if (result.length === 0) {
                            if (element.FieldTypeKind) {
                                let fieldCollection = spWeb.lists.getById(listId).fields;
                                if (element.FieldTypeKind === FieldTypeKind.Lookup) {
                                    this.addLookupFieldToCollection(fieldCollection, element).then(() => {
                                        resolve();
                                    }).catch((error) => {
                                        reject(error);
                                    });
                                }
                                else if (element.FieldTypeKind === FieldTypeKind.Calculated) {
                                    this.addCalcFieldToCollection(fieldCollection, element).then(() => {
                                        resolve();
                                    }).catch((error) => {
                                        reject(error);
                                    });
                                } else {
                                    this.addFieldToCollection(fieldCollection, element).then(() => {
                                        resolve();
                                    }).catch((error) => {
                                        reject(error);
                                    });
                                }
                            } else {
                                let error = `FieldTypKind for '${element.InternalName}' could not be resolved`;
                                reject(error);
                            }
                        } else {
                            resolve();
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

    private addFieldSite(url: string, pConfig: IField): Promise<any> {
        return new Promise((resolve, reject) => {
            let spWeb = new Web(url);
            let element = pConfig;
            spWeb.fields.filter("InternalName eq '" + element.InternalName + "'").select("Id").get().then((data) => {
                if (data.length === 0) {
                    if (element.FieldTypeKind) {
                        let fieldCollection = spWeb.fields;
                        if (element.FieldTypeKind === FieldTypeKind.Lookup) {
                            this.addLookupFieldToCollection(fieldCollection, element).then(() => {
                                resolve();
                            }).catch((error) => {
                                reject(error);
                            });
                        }
                        else if (element.FieldTypeKind === FieldTypeKind.Calculated) {
                            this.addCalcFieldToCollection(fieldCollection, element).then(() => {
                                resolve();
                            }).catch((error) => {
                                reject(error);
                            });
                        }
                        else {
                            this.addFieldToCollection(fieldCollection, element).then(() => {
                                resolve();
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
                    resolve();
                    Logger.write(`Field with Internal Name '${element.InternalName}' already exists`);
                }
            }).catch((error) => {
                reject(error);
            });
        });
    }


    private addFieldToCollection(pFieldCol: Fields, element: IField) {
        return new Promise((resolve, reject) => {
            let propertyHash = this.CreateProperties(element);
            pFieldCol.add(element.InternalName, "SP.Field", propertyHash).then((result) => {
                result.field.update({ Title: element.Title }).then(() => {
                    Logger.write(`Field with Internal Name '${element.InternalName}' created`);
                    resolve();
                }).catch((error) => {
                    reject(error);
                });
            }).catch((error) => {
                reject(error);
            });
        });
    }

    private addCalcFieldToCollection(pFieldCol: Fields, element: IField) {
        return new Promise((resolve, reject) => {
            let propertyHash = this.CreateProperties(element);
            pFieldCol.addCalculated(element.InternalName, element.Formula, Types.DateTimeFieldFormatType[element.DateFormat], Types.FieldTypes[element.OutputType], propertyHash).then((result) => {
                result.field.update({ Title: element.Title, Description: element.Description }).then(() => {
                    resolve();
                    Logger.write("Calculated Field with Internal Name '" + element.InternalName + "' created", 1);
                }).catch((error) => {
                    reject(error);
                });
            }).catch((error) => {
                reject(error);
            });
        });
    }

    private addLookupFieldToCollection(pFieldCol: Fields, element: IField) {
        return new Promise((resolve, reject) => {
            resolve();
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

