import * as Types from "@agileis/sp-pnp-js/lib/sharepoint/rest/types";
import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Fields } from "@agileis/sp-pnp-js/lib/sharepoint/rest/Fields";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { Field } from "@agileis/sp-pnp-js/lib/sharepoint/rest/fields";
import { List } from "@agileis/sp-pnp-js/lib/sharepoint/rest/Lists";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { IField }  from "../Interfaces/Types/IField";
import { FieldTypeKind } from "../Constants/FieldTypeKind";
import { ControlOption } from "../Constants/ControlOption";
import { Util } from "../Util/Util";

class LookupFieldInfo {
    public clientContext: SP.ClientContext;
    public lookupList: SP.List;
    public fieldCollection: SP.FieldCollection;
    constructor(clientContext: SP.ClientContext, lookupList: SP.List, fieldCollection: SP.FieldCollection) {
        this.clientContext = clientContext;
        this.lookupList = lookupList;
        this.fieldCollection = fieldCollection;
    }
}

export class FieldHandler implements ISPObjectHandler {
    public execute(fieldConfig: IField, parentPromise: Promise<IPromiseResult<List | Web>>): Promise<IPromiseResult<void | Field>> {
        return new Promise<IPromiseResult<void | Field>>((resolve, reject) => {
            parentPromise.then((promiseResult) => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, fieldConfig.InternalName,
                        `Field handler parent promise value result is null or undefined for the field with the internal name '${fieldConfig.InternalName}'!`);
                } else {
                    this.processingFieldConfig(fieldConfig, promiseResult.value.fields)
                        .then((fieldProcessingResult) => { resolve(fieldProcessingResult); })
                        .catch((error) => {
                            Util.Retry(error, fieldConfig.InternalName,
                                () => {
                                    return this.processingFieldConfig(fieldConfig, promiseResult.value.fields);
                                });
                        });
                }
            });
        });
    }

    private processingFieldConfig(fieldConfig: IField, fieldCollection: Fields): Promise<IPromiseResult<void | Field>> {
        return new Promise<IPromiseResult<void | Field>>((resolve, reject) => {
            let processingText = fieldConfig.ControlOption === ControlOption.Add || fieldConfig.ControlOption === undefined || fieldConfig.ControlOption === ""
                ? "Add" : fieldConfig.ControlOption;
            Logger.write(`Processing ${processingText} field: '${fieldConfig.Title}'.`, Logger.LogLevel.Info);

            fieldCollection.filter(`InternalName eq '${fieldConfig.InternalName}'`).select("Id").get()
                .then((fieldRequestResults) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void | Field>> = undefined;

                    if (fieldRequestResults && fieldRequestResults.length === 1) {
                        Logger.write(`Found Field with the internal name: '${fieldConfig.InternalName}'`);
                        let field = fieldCollection.getById(fieldRequestResults[0].Id);
                        switch (fieldConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateField(fieldConfig, field);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteField(fieldConfig, field);
                                break;
                            default:
                                fieldConfig.ControlOption = ControlOption.Update;
                                Util.Resolve<Field>(resolve, fieldConfig.InternalName, `Field with the internal name '${fieldConfig.InternalName}'` +
                                    ` does not have to be added, because it already exists.`, field);
                                rejectOrResolved = true;
                                break;
                        }
                    } else {
                        switch (fieldConfig.ControlOption) {
                            case ControlOption.Delete:
                                Util.Resolve<void>(resolve, fieldConfig.InternalName, `Field with internal name '${fieldConfig.InternalName}' does not have to be deleted, because it does not exist.`);
                                rejectOrResolved = true;
                                break;
                            case ControlOption.Update:
                                fieldConfig.ControlOption = ControlOption.Add;
                            default:
                                processingPromise = this.addField(fieldConfig, fieldCollection);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((fieldProcessingResult) => { resolve(fieldProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write("Field handler processing promise is undefined!", Logger.LogLevel.Error);
                    }
                })
                .catch((error) => { Util.Reject<void>(reject, fieldConfig.InternalName, `Error while requesting field with the internal name '${fieldConfig.InternalName}': ` + error); });
        });
    }

    private addField(fieldConfig: IField, fieldCollection: Fields): Promise<IPromiseResult<Field>> {
        return new Promise<IPromiseResult<Field>>((resolve, reject) => {
            let processingPromise: Promise<IPromiseResult<Field>> = undefined;

            switch (FieldTypeKind[fieldConfig.FieldType]) {
                case undefined:
                case "":
                case null:
                    Util.Reject<void>(reject, fieldConfig.InternalName, `Field type kind could not be resolved for the field with the internal name ${fieldConfig.InternalName}`);
                    break;
                case FieldTypeKind.Lookup:
                    processingPromise = this.addLookupField(fieldConfig, fieldCollection);
                    break;
                case FieldTypeKind.Calculated:
                    processingPromise = this.addCalculatedField(fieldConfig, fieldCollection);
                    break;
                default:
                    processingPromise = this.addDefaultField(fieldConfig, fieldCollection);
                    break;
            }

            if (processingPromise) {
                processingPromise
                    .then((fieldProcessingResult) => { resolve(fieldProcessingResult); })
                    .catch((error) => { reject(error); });
            }
        });
    }

    private addDefaultField(fieldConfig: IField, fieldCollection: Fields): Promise<IPromiseResult<Field>> {
        return new Promise<IPromiseResult<Field>>((resolve, reject) => {
            let properties = this.createProperties(fieldConfig);
            fieldCollection.add(fieldConfig.InternalName, "SP.Field", properties)
                .then((fieldAddResult) => {
                    fieldConfig.ControlOption = ControlOption.Update;
                    this.updateField(fieldConfig, fieldAddResult.field)
                        .then((fieldUpdateResult) => {
                            Util.Resolve<Field>(resolve, fieldConfig.InternalName, `Added field: '${fieldConfig.InternalName}'.`, fieldUpdateResult.value);
                        })
                        .catch((error) => {
                            Util.Reject<void>(reject, fieldConfig.InternalName,
                                `Error while adding and updating field with the internal name '${fieldConfig.InternalName}': ` + error);
                        });
                })
                .catch((error) => {
                    this.tryToDeleteCorruptedField(fieldConfig, fieldCollection)
                        .then(() => { Util.Reject<void>(reject, fieldConfig.InternalName, `Error while adding field with the internal name '${fieldConfig.InternalName}' - field deleted: ` + error); })
                        .catch(() => {
                            Util.Reject<void>(reject, fieldConfig.InternalName,
                                `Error while adding field with the internal name '${fieldConfig.InternalName}' - field not deleted: ` + error);
                        });
                });
        });
    }

    private addCalculatedField(fieldConfig: IField, fieldCollection: Fields): Promise<IPromiseResult<Field>> {
        return new Promise<IPromiseResult<Field>>((resolve, reject) => {
            let properties = this.createProperties(fieldConfig);
            fieldCollection.addCalculated(fieldConfig.InternalName, fieldConfig.Formula, Types.DateTimeFieldFormatType[fieldConfig.DateFormat], Types.FieldTypes[fieldConfig.OutputType], properties)
                .then((fieldAddResult) => {
                    fieldConfig.ControlOption = ControlOption.Update;
                    this.updateField(fieldConfig, fieldAddResult.field)
                        .then((fieldUpdateResult) => {
                            Util.Resolve<Field>(resolve, fieldConfig.InternalName, `Added field: '${fieldConfig.InternalName}'.`, fieldUpdateResult.value);
                        })
                        .catch((error) => {
                            Util.Reject<void>(reject, fieldConfig.InternalName,
                                `Error while adding and updating field with the internal name '${fieldConfig.InternalName}': ` + error);
                        });
                })
                .catch((error) => {
                    this.tryToDeleteCorruptedField(fieldConfig, fieldCollection)
                        .then(() => { Util.Reject<void>(reject, fieldConfig.InternalName, `Error while adding field with the internal name '${fieldConfig.InternalName}' - field deleted: ` + error); })
                        .catch(() => {
                            Util.Reject<void>(reject, fieldConfig.InternalName,
                                `Error while adding field with the internal name '${fieldConfig.InternalName}' - field not deleted: ` + error);
                        });
                });
        });
    }

    private getLookupFieldInfo(fieldConfig: IField, fieldCollection: Fields): Promise<IPromiseResult<LookupFieldInfo>> {
        return new Promise<IPromiseResult<LookupFieldInfo>>((resolve, reject) => {
            let context = SP.ClientContext.get_current();
            let lookupList: SP.List = context.get_web().get_lists().getByTitle(fieldConfig.LookupList);

            let urlParts = fieldCollection.toUrl().split("/").reverse();
            let spFieldCollection: SP.FieldCollection = undefined;
            if ((urlParts[1] as string).indexOf("lists") === 0) {
                let listId = (urlParts[1] as string).split("'")[1];
                spFieldCollection = context.get_web().get_lists().getById(listId).get_fields();
            } else {
                spFieldCollection = context.get_web().get_fields();
            }

            context.load(lookupList);
            context.executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<LookupFieldInfo>(resolve, undefined, undefined, new LookupFieldInfo(context, lookupList, spFieldCollection));
                },
                (sender, args) => {
                    Util.Reject<void>(reject, fieldConfig.InternalName,
                        `Error while requesting lookup list and lookup field collection in adding lookup field with internal name '${fieldConfig.InternalName}':
                            ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                });
        });
    }

    private updateLookupFieldProperties(fieldConfig: IField, lookupField: SP.FieldLookup) {
        for (let prop in fieldConfig) {
            switch (prop) {
                case "Title":
                    lookupField.set_title(fieldConfig[prop]);
                    break;
                case "Required":
                    lookupField.set_required(fieldConfig[prop]);
                    break;
                case "Indexed":
                    lookupField.set_indexed(fieldConfig[prop]);
                    break;
                case "JSLink":
                    lookupField.set_jsLink(fieldConfig[prop]);
                    break;
                case "RelationshipDeleteBehavior":
                    lookupField.set_relationshipDeleteBehavior(SP.RelationshipDeleteBehaviorType[(fieldConfig[prop] as string).toLowerCase()]);
                    break;
            }
        }

        lookupField.update();
    }

    private addLookupField(fieldConfig: IField, fieldCollection: Fields): Promise<IPromiseResult<Field>> {
        return new Promise<IPromiseResult<Field>>((resolve, reject) => {
            this.getLookupFieldInfo(fieldConfig, fieldCollection)
                .then((lookupInfoResult) => {
                    let context = lookupInfoResult.value.clientContext;
                    let lookupList = lookupInfoResult.value.lookupList;
                    let spFieldCollection = lookupInfoResult.value.fieldCollection;

                    const fieldXml = `<Field Type='${fieldConfig.FieldType}' ${fieldConfig.Multivalue ? "Mult='TRUE'" : ""} DisplayName='${fieldConfig.InternalName}'` +
                        ` ShowField='${fieldConfig.LookupField}' List='${lookupList.get_id().toString()}' Name='${fieldConfig.InternalName}'></Field>`;

                    let lookupField = spFieldCollection.addFieldAsXml(fieldXml, false, SP.AddFieldOptions.addToDefaultContentType);
                    this.updateLookupFieldProperties(fieldConfig, <SP.FieldLookup>context.castTo(lookupField, SP.FieldLookup));

                    fieldConfig.DependendFields.forEach(dependendField => {
                        let depFieldInternalName = `${lookupList.get_title()}_${dependendField.InternalName}`.substr(0, 32);
                        let depField = spFieldCollection.addDependentLookup(depFieldInternalName, lookupField, dependendField.InternalName);
                        depField.set_title(`${lookupList.get_title()}:${dependendField.Title}`);
                        depField.update();
                    });

                    context.load(lookupField, "Id");
                    context.executeQueryAsync(
                        (sender, args) => {
                            Util.Resolve<Field>(resolve, fieldConfig.InternalName, `Added field: '${fieldConfig.InternalName}'.`, fieldCollection.getById(lookupField.get_id().toString()));
                        },
                        (sender, args) => {
                            Util.Reject<void>(reject, fieldConfig.InternalName, `Error while adding and updating lookup field with internal name '${fieldConfig.InternalName}': ${args.get_message()}'`
                                + `\n' ${args.get_stackTrace()}`);
                        });
                })
                .catch((error) => reject(error));
        });
    }

    private updateField(fieldConfig: IField, field: Field): Promise<IPromiseResult<Field>> {
        return new Promise<IPromiseResult<Field>>((resolve, reject) => {
            let properties = this.createProperties(fieldConfig);

            let type = `SP.Field${fieldConfig.FieldType ? fieldConfig.FieldType : ""}`;
            switch (fieldConfig.FieldType) {
                case "Boolean":
                    type = "SP.Field";
                    break;
                case "Note":
                    type = "SP.FieldMultiLineText";
                    break;
            }

            field.update(properties, type)
                .then((fieldUpdateResult) => {
                    Util.Resolve<Field>(resolve, fieldConfig.InternalName, `Updated field: '${fieldConfig.InternalName}'.`,
                        fieldUpdateResult.field);
                })
                .catch((error) => { Util.Reject<void>(reject, fieldConfig.InternalName, `Error while updating field with internal name '${fieldConfig.InternalName}': ` + error); });
        });
    }

    private deleteField(fieldConfig: IField, field: Field): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            field.delete()
                .then(() => { Util.Resolve<void>(resolve, fieldConfig.InternalName, `Deleted field: '${fieldConfig.InternalName}'.`); })
                .catch((error) => {
                    Util.Reject<void>(reject, fieldConfig.InternalName,
                        `Error while deleting field with internal name '${fieldConfig.InternalName}': ` + error);
                });
        });
    }

    private tryToDeleteCorruptedField(fieldConfig: IField, fields: Fields): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            fields.getByInternalNameOrTitle(fieldConfig.InternalName).delete()
                .then(() => { resolve(); })
                .catch((error) => { reject(error); });
        });
    }

    private createProperties(fieldConfig: IField) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(fieldConfig);
        let parsedObject = JSON.parse(stringifiedObject);
        switch (fieldConfig.ControlOption) {
            case ControlOption.Update:
                delete parsedObject.ControlOption;
                delete parsedObject.InternalName;
                delete parsedObject.FieldType;
                delete parsedObject.DateFormat;
                delete parsedObject.OutputType;
                delete parsedObject.DependendFields;
                delete parsedObject.LookupList;
                if (fieldConfig.RelationshipDeleteBehavior) {
                    parsedObject.RelationshipDeleteBehavior = SP.RelationshipDeleteBehaviorType[fieldConfig.RelationshipDeleteBehavior.toLowerCase()];
                }
                if (fieldConfig.DisplayFormat) {
                    parsedObject.DisplayFormat = Types.UrlFieldFormatType[fieldConfig.DisplayFormat];
                }
                break;
            default:
                delete parsedObject.ControlOption;
                delete parsedObject.Title;
                delete parsedObject.InternalName;
                delete parsedObject.DateFormat;
                delete parsedObject.Formula;
                delete parsedObject.OutputType;
                delete parsedObject.NumberOfLines;
                delete parsedObject.AppendOnly;
                delete parsedObject.DisplayFormat;
                parsedObject.FieldTypeKind = FieldTypeKind[parsedObject.FieldType];
                delete parsedObject.FieldType;
                if (fieldConfig.DisplayFormat) {
                    parsedObject.DisplayFormat = Types.UrlFieldFormatType[fieldConfig.DisplayFormat];
                }
                break;
        }

        if (FieldTypeKind[fieldConfig.FieldType] === FieldTypeKind.DateTime && parsedObject.DisplayFormat) {
            switch (parsedObject.DisplayFormat) {
                case "DateOnly":
                    parsedObject.DisplayFormat = 0;
                    break;
                default:
                    parsedObject.DisplayFormat = 1;
                    break;
            }
        }

        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}

