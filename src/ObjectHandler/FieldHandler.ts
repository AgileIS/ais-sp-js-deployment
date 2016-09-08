import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IField}  from "../interface/Types/IField";
import {IList} from "../interface/Types/IList";
import {ISite} from "../interface/Types/ISite";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import {Fields} from "sp-pnp-js/lib/sharepoint/rest/Fields";
import {Web} from "sp-pnp-js/lib/sharepoint/rest/webs";
import {Field} from "sp-pnp-js/lib/sharepoint/rest/fields";
import {List} from "sp-pnp-js/lib/sharepoint/rest/Lists";
import {Reject, Resolve} from "../Util/Util";
import {FieldTypeKind} from "../Constants/FieldTypeKind";

export class FieldHandler {

    execute(config: IField, parent: Promise<List | Web>) {

        Logger.write("config " + JSON.stringify(config));
        switch (config.ControlOption) {
            case "Update":
                return this.updateField(config, parent);
            case "Delete":
                return this.deleteField(config, parent);
            default:
                return this.addField(config, parent);
        }
    }


    private addField(config: IField, parent: Promise<Web | List>): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            parent.then(parentInstance => {
                Logger.write("config " + JSON.stringify(config), 0);
                Logger.write("Entering add Field", 1);
                parentInstance.fields.filter(`InternalName eq '${config.InternalName}'`).select("Id").get().then((result) => {
                    if (result.length === 0) {
                        if (config.FieldTypeKind) {
                            if (config.FieldTypeKind === FieldTypeKind.Lookup) {
                                let error = `Not implemented yet - createFieldAsXml`;
                                 Reject(reject, error, config.Title);
                            }
                            else if (config.FieldTypeKind === FieldTypeKind.Calculated) {
                                let propertyHash = this.CreateProperties(config);
                                parentInstance.fields.addCalculated(config.InternalName, config.Formula, Types.DateTimeFieldFormatType[config.DateFormat], Types.FieldTypes[config.OutputType], propertyHash).then((result) => {
                                    result.field.update({ Title: config.Title, Description: config.Description }).then((result) => {
                                        let field = parentInstance.fields.getById(result.data.Id);
                                        resolve(field);
                                        Logger.write("Calculated Field with Internal Name '" + config.InternalName + "' created", 1);
                                    }).catch((error) => {
                                         Reject(reject, error, config.Title);
                                    });
                                }).catch((error) => {
                                     Reject(reject, error, config.Title);
                                });
                            } else {
                                let propertyHash = this.CreateProperties(config);
                                parentInstance.fields.add(config.InternalName, "SP.Field", propertyHash).then((result) => {
                                    result.field.update({ Title: config.Title }).then(() => {
                                        let field = parentInstance.fields.getById(result.data.Id);
                                        Logger.write(`Field with Internal Name '${config.InternalName}' created`);
                                        resolve(field);
                                    }).catch((error) => {
                                         Reject(reject, error, config.Title);
                                    });
                                }).catch((error) => {
                                     Reject(reject, error, config.Title);
                                });
                            }
                        } else {
                            let error = `FieldTypKind for '${config.InternalName}' could not be resolved`;
                             Reject(reject, error, config.Title);
                        }
                    } else {
                        let field = parentInstance.fields.getById(result[0].Id);
                        resolve(field);
                        Logger.write(`Field with Internal Name '${config.InternalName}' already exists`);
                    }
                }).catch((error) => {
                     Reject(reject, error, config.Title);
                });
            });
        });
    }


    private updateField(config: IField, parent: Promise<Web | List>): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            parent.then(parentInstance => {
                Logger.write("config " + JSON.stringify(config), 0);
                Logger.write("Entering update Field", 1);
                parentInstance.fields.filter(`InternalName eq '${config.InternalName}'`).select("Id").get().then((result) => {
                    if (result.length === 1) {
                        let field =  parentInstance.fields.getById(result[0].Id);
                        let properties = this.CreateProperties(config);
                        field.update(properties).then(() => {
                            resolve(field);
                            Logger.write(`Field with Internal Name '${config.InternalName}' updated`, 1);
                        }).catch((error) => {
                             Reject(reject, error, config.Title);
                        });
                    }
                    else {
                        let error = `Field with Internal Name '${config.InternalName}' does not exist`;
                         Reject(reject, error, config.Title);
                    }
                });
            });
        });
    }



    private deleteField(config: IField, parent: Promise<Web | List>): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            parent.then(parentInstance => {
                Logger.write("config " + JSON.stringify(config), 0);
                Logger.write("Entering delete Field", 1);
                parentInstance.fields.filter(`InternalName eq '${config.InternalName}'`).select("Id").get().then((result) => {
                    if (result.length === 1) {
                        let field = parentInstance.fields.getById(result[0].Id);
                        field.delete().then(() => {
                            Logger.write(`Field with Internal Name '${config.InternalName}' deleted`, 1);
                            resolve(field);
                        }).catch((error) => {
                             Reject(reject, error, config.Title);
                        });
                    }
                    else {
                        let error = `Field with Internal Name '${config.InternalName}' does not exist`;
                         Reject(reject, error, config.Title);
                    }
                });
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

