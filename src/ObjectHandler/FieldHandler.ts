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
        return new Promise<Field>((resolve, reject) => {
            parent.then(parentInstance => {
                Logger.write("config " + JSON.stringify(config));
                let fieldprocessed: Array<Promise<any>> = [];
                switch (config.ControlOption) {
                    case "Update":
                        fieldprocessed.push(this.updateField(config, parentInstance));
                        break;
                    case "Delete":
                        fieldprocessed.push(this.deleteField(config, parentInstance));
                        break;
                    default:
                        fieldprocessed.push(this.createField(config, parentInstance));
                        fieldprocessed.push(this.updateField(config, parentInstance));
                }

                Promise.all(fieldprocessed).then(result => {
                    Resolve(resolve, `Field with Internal Name '${config.InternalName}' processed`, config.InternalName)
                }).catch(error => {
                    Reject(reject, error, config.InternalName);
                });

            });
        });
    }


    private createField(config: IField, parentInstance: Web | List): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config), 0);
            Logger.write("Entering add Field", 1);
            parentInstance.fields.filter(`InternalName eq '${config.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 0) {
                    if (config.FieldTypeKind) {
                        let field = this.addField(config, parentInstance);
                        Resolve(resolve, "Field created", config.InternalName, field);
                    } else {
                        let error = `FieldTypKind for '${config.InternalName}' could not be resolved`;
                        Reject(reject, error, config.Title);
                    }
                } else {
                    let field = parentInstance.fields.getById(result[0].Id);
                    Resolve(resolve, `Field with InternalName '${config.InternalName}' already exists`, config.InternalName, field);
                }
            }).catch((error) => {
                Reject(reject, error, config.Title);
            });
        });

    }


    private updateField(config: IField, parentInstance: Web | List): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config), 0);
            Logger.write("Entering update Field", 1);
            parentInstance.fields.filter(`InternalName eq '${config.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let listId = result[0].Id;
                    let field = parentInstance.fields.getById(result[0].Id);
                    let properties = this.CreateProperties(config);
                    field.update(properties).then((result) => {
                        let fieldAfterUpdate = parentInstance.fields.getById(listId);
                        Resolve(resolve, `Field with Internal Name '${config.InternalName}' updated`, config.InternalName, fieldAfterUpdate);
                    }).catch((error) => {
                        Reject(reject, error, config.Title);
                    });
                }
                else {
                    let error = `Field with Internal Name '${config.InternalName}' does not exist`;
                    Reject(reject, error, config.Title);
                }
            }).catch((error) => {
                Reject(reject, error, config.Title);
            });
        });
    }



    private deleteField(config: IField, parentInstance: Web | List): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config), 0);
            Logger.write("Entering delete Field", 1);
            parentInstance.fields.filter(`InternalName eq '${config.InternalName}'`).select("Id").get().then((result) => {
                if (result.length === 1) {
                    let field = parentInstance.fields.getById(result[0].Id);
                    field.delete().then(() => {
                        Resolve(resolve, `Field with Internal Name '${config.InternalName}' deleted`, config.InternalName, field);
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
    }


    private addField(config: IField, parentInstance: Web | List): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            if (config.FieldTypeKind === FieldTypeKind.Lookup) {
                let error = `Not implemented yet - createFieldAsXml`;
                Reject(reject, error, config.Title);
            }
            else if (config.FieldTypeKind === FieldTypeKind.Calculated) {
                let propertyHash = this.CreateProperties(config);
                parentInstance.fields.addCalculated(config.InternalName, config.Formula, Types.DateTimeFieldFormatType[config.DateFormat], Types.FieldTypes[config.OutputType], propertyHash).then(() => {
                    this.updateField(config, parentInstance).then((result) => {
                        let field = result;
                        Resolve(resolve, `Calculated Field with Internal Name ' ${config.InternalName}' created`, config.InternalName, field);
                    }).catch((error) => {
                        Reject(reject, error, config.Title);
                    });
                }).catch((error) => {
                    Reject(reject, error, config.Title);
                });
            } else {
                let propertyHash = this.CreateProperties(config);
                parentInstance.fields.add(config.InternalName, "SP.Field", propertyHash).then(() => {
                    this.updateField(config, parentInstance).then((result) => {
                        let field = result;
                        Resolve(resolve, `Field with Internal Name '${config.InternalName}' created`, config.InternalName, field);
                    }).catch((error) => {
                        Reject(reject, error, config.Title);
                    });
                }).catch((error) => {
                    Reject(reject, error, config.Title);
                });
            }
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
                delete parsedObject.InternalName;
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

