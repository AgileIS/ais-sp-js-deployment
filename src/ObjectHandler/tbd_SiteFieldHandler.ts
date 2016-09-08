import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IField} from "../interface/Types/IField";
import {IList} from "../interface/Types/IList";
import {ISite} from "../interface/Types/ISite";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";
import {RejectAndLog} from "../lib/Util/Util";
import {FieldTypeKind} from "../lib/FieldTypeKind";

export class SiteFieldHandler  {
    public execute(config: IField, url: string, parentConfig: ISite) {
        switch (config.ControlOption) {
            case "":
                return this.AddField(config, url);
            case "Update":
                return this.UpdateField(config, url);
            case "Delete":
                return this.DeleteField(config, url);
            default:
                return this.AddField(config, url);
        }
    }

    private AddField(config: IField, url: string) {
        let spWeb = new web.Web(url);
        let element = config;
        return new Promise<IField>((resolve, reject) => {
            Logger.write("config " + JSON.stringify(config));
            spWeb.fields.filter("InternalName eq '" + element.InternalName + "'").select("Id").get().then((data) => {
                if (data.length === 0) {
                    if (element.FieldTypeKind) {
                        if (element.FieldTypeKind === FieldTypeKind.Lookup) {
                            resolve(element);
                        }
                        else if (element.FieldTypeKind === FieldTypeKind.Calculated) {
                            let propertyHash = this.CreateProperties(element);
                            spWeb.fields.addCalculated(element.InternalName, element.Formula, Types.DateTimeFieldFormatType[element.DateFormat], Types.FieldTypes[element.OutputType], propertyHash).then((result) => {
                                result.field.update({ Title: element.Title, Description: element.Description }).then(() => {
                                    resolve(element);
                                    Logger.write("Calculated Field with Internal Name '" + element.InternalName + "' created", 1);
                                }).catch((error) => {
                                    RejectAndLog(error, element.InternalName, reject);
                                });
                            }).catch((error) => {
                                RejectAndLog(error, element.InternalName, reject);
                            });
                        }
                        else {
                            let propertyHash = this.CreateProperties(element);
                            spWeb.fields.add(element.InternalName, "SP.Field", propertyHash).then((result) => {
                                result.field.update({ Title: element.Title }).then(() => {
                                    resolve(element);
                                    Logger.write("Field with Internal Name'" + element.InternalName + "' created", 1);
                                }).catch((error) => {
                                    RejectAndLog(error, element.InternalName, reject);
                                });
                            }).catch((error) => {
                                RejectAndLog(error, element.InternalName, reject);
                            });
                        }
                    }
                    else {
                        let error = "FieldTypKind could not be resolved";
                        RejectAndLog(error, element.InternalName, reject);
                    }
                }
                else {
                    let error = "Field with Internal Name '" + element.InternalName + "' already exists";
                    RejectAndLog(error, element.InternalName, reject);
                }
            }).catch((error) => {
                RejectAndLog(error, element.InternalName, reject);
            });
        });
    };


    private UpdateField(config: IField, url: string) {
        let spWeb = new web.Web(url);
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
        let spWeb = new web.Web(url);
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