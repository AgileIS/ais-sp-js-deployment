import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IField} from "../interface/Types/IField";
import {IList} from "../interface/Types/IList";
import {ISite} from "../interface/Types/ISite";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";

export class FieldHandler implements ISPObjectHandler {
    execute(config: IField, url: string, parentConfig: IList) {
        let spWeb = new web.Web(url);
        let element = config;
        let parentElement = parentConfig;
        return new Promise<IField>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config));
            spWeb.lists.filter(`RootFolder/Name eq '${parentElement.InternalName}'`).select("Id").get().then((data) => {
                if (data.length === 1) {
                    let listId = data[0].Id;
                    spWeb.lists.getById(listId).fields.filter(`InternalName eq '${element.InternalName}'`).select("Id").get().then((result) => {
                        if (result.length === 0) {
                            if (element.FieldTypeKind) {
                                if (element.FieldTypeKind === "Lookup") {
                                    // Lookup nicht vorhanden umsetzten als create FieldAsXml
                                }
                                else if (element.FieldTypeKind === "Calculated") {
                                    spWeb.lists.getById(listId).fields.addCalculated(element.InternalName, element.Formula, Types.DateTimeFieldFormatType.DateOnly).then((result) => {
                                        result.field.update({ Title: element.Title }).then(
                                            () => {
                                                Logger.write(`Calculated Field with Internal Name '${element.InternalName}' created`);
                                                resolve(config);
                                            },
                                            (error) => {
                                                reject(error);
                                            }
                                        );
                                    },
                                    (error) => {
                                        reject(error);
                                    });
                                }
                                else {
                                    let propertyHash = createTypedHashfromProperties(element);
                                    spWeb.lists.getById(listId).fields.add(element.InternalName, "SP.Field", propertyHash).then((result) => {
                                        result.field.update({ Title: element.Title }).then(
                                            () => {
                                                Logger.write(`Field with Internal Name '${element.InternalName}' created`);
                                                resolve(config);
                                            },
                                            (error) => {
                                                reject(error);
                                            }
                                        );
                                    },
                                    (error) => {
                                        reject(error);
                                    }
                                    );
                                }
                            }
                            else {
                                let error = `FieldTypKind for '${element.InternalName}' could not be resolved`;
                                reject(error);
                                Logger.write(error, 0);
                            }
                        }
                        else if (result.length === 1 && element.FieldTypeKind === undefined) {
                            let fieldId = result[0].Id;
                            spWeb.lists.getById(listId).fields.getById(fieldId).update({ Title: element.Title }).then(
                                () => {
                                    Logger.write("Existing Field with Title '" + element.Title + "' updated");
                                    resolve(config);
                                },
                                (error) => {
                                    reject(error);
                                }
                            );
                        }
                        else {
                            resolve(config);
                            Logger.write(`Field with Internal Name '${element.InternalName}' already exists`);
                        }
                    });
                }
                else {
                    let error = `List with Title '${parentElement.InternalName}' for Field '${element.InternalName}' does not exist`;
                    reject(error);
                    Logger.write(error, 0);
                }
            })
        });
    }
}

function createFieldHash(pElement: IField) {
    let element = pElement;
    let stringifiedObject = JSON.stringify(element);
    return JSON.parse(stringifiedObject);
}

function createTypedHashfromProperties(pElement: IField) {
    let element = pElement;
    let stringifiedObject: string;
    stringifiedObject = JSON.stringify(element);
    let parsedObject = JSON.parse(stringifiedObject);
    delete parsedObject.ControlOption;
    delete parsedObject.Title;
    delete parsedObject.Description;
    stringifiedObject = JSON.stringify(parsedObject);
    return JSON.parse(stringifiedObject);
}