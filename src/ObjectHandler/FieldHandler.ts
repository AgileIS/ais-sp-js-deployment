import {ISPObjectHandler} from "../interface/ObjectHandler/ispobjecthandler";
import {Logger} from "sp-pnp-js/lib/utils/logging";
import {IField} from "../interface/Types/IField";
import {IList} from "../interface/Types/IList";
import {ISite} from "../interface/Types/ISite";
import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";

export class FieldHandler implements ISPObjectHandler {
    execute(config: IField, url: string, parent: Promise<ISite | IList>) {
        let spWeb = new web.Web(url);
        let element = config;
        let parentPromise = parent;
        return new Promise<IField>((resolve, reject) => {
            parentPromise.then(() => {
                Logger.write("config " + JSON.stringify(config));
                spWeb.fields.filter("InternalName eq '" + element.InternalName + "'").get().then(function (data) {
                    if (data.length === 0) {
                        if (element.FieldTypeKind) {
                            if (element.FieldTypeKind === 7) {  // 7 = Lookup
                                resolve();
                            }
                            else if (element.FieldTypeKind === 17) { // 17 = Calculated
                                spWeb.fields.addCalculated(element.InternalName, element.Formula, Types.DateTimeFieldFormatType.DateOnly).then(function (result) {
                                    result.field.update({ Title: element.Title }).then(function () {
                                        resolve();
                                        Logger.write("Calculated Field with Internal Name '" + element.InternalName + "' created");
                                    });
                                });
                            }
                            else {
                                let propertyHash = createTypedHashfromProperties(element);
                                spWeb.fields.add(element.InternalName, "SP.Field", propertyHash).then(function (result) {
                                    result.field.update({ Title: element.Title }).then(() => {
                                        resolve();
                                        Logger.write("Field with Internal Name'" + element.InternalName + "' created");
                                    });
                                });
                            }
                        }
                        else {
                            let error = "FieldTypKind could not be resolved";
                            reject(error);
                            Logger.write(error);
                        }
                    }
                    else {
                        let error = "Field with Internal Name '" + element.InternalName + "' already exists";
                        reject(error);
                        Logger.write(error);
                    }
                });
            }, () => {
                let error = "Parent Promise not resolved";
                reject(error);
                Logger.write(error);
            });
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
    delete parsedObject.Title;
    stringifiedObject = JSON.stringify(parsedObject);
    return JSON.parse(stringifiedObject);
}