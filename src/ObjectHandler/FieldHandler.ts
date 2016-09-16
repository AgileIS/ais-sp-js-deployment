import * as Types from "sp-pnp-js/lib/sharepoint/rest/types";
import { Logger } from "sp-pnp-js/lib/utils/logging";
import { Fields } from "sp-pnp-js/lib/sharepoint/rest/Fields";
import { Web } from "sp-pnp-js/lib/sharepoint/rest/webs";
import { Field } from "sp-pnp-js/lib/sharepoint/rest/fields";
import { List } from "sp-pnp-js/lib/sharepoint/rest/Lists";
import { IField }  from "../interface/Types/IField";
import { FieldTypeKind } from "../Constants/FieldTypeKind";
import { ControlOption } from "../Constants/ControlOption";
import { Reject, Resolve } from "../Util/Util";


export class FieldHandler {
    public execute(fieldConfig: IField, parentPromise: Promise<List | Web>): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            parentPromise.then((parentInstance) => {
                this.ProcessingViewConfig(fieldConfig, parentInstance.fields)
                    .then((viewProsssingResult) => { resolve(viewProsssingResult); })
                    .catch((error) => { reject(error); });
            });
        });
    }

    private ProcessingViewConfig(fieldConfig: IField, parentFields: Fields): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            Logger.write(`Processing ${fieldConfig.ControlOption === ControlOption.Add || fieldConfig.ControlOption === undefined
                ? "Add" : fieldConfig.ControlOption} field: '${fieldConfig.Title}'`, Logger.LogLevel.Info);

            parentFields.filter(`InternalName eq '${fieldConfig.InternalName}'`).select("Id").get()
                .then((fieldRequestResults) => {
                    let processingPromise: Promise<Field> = undefined;

                    if (fieldRequestResults && fieldRequestResults.length === 1) {
                        let field = parentFields.getById(fieldRequestResults[0].Id);
                        switch (fieldConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateField(fieldConfig, field);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteField(fieldConfig, field);
                                break;
                            default:
                                Resolve(resolve, `Field with the internal name '${fieldConfig.InternalName}' already exists`, fieldConfig.Title, field);
                                break;
                        }
                    } else {
                        switch (fieldConfig.ControlOption) {
                            case ControlOption.Update:
                            case ControlOption.Delete:
                                Reject(reject, `field with internal name '${fieldConfig.InternalName}' does not exists`, fieldConfig.Title);
                                break;
                            default:
                                processingPromise = this.addField(fieldConfig, parentFields);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((fieldProcessingResult) => { resolve(fieldProcessingResult); })
                            .catch((error) => { reject(error); });
                    }
                })
                .catch((error) => { Reject(reject, `Error while requesting field with the internal name '${fieldConfig.InternalName}': ` + error, fieldConfig.Title); });
        });
    }

    private addField(fieldConfig: IField, parentFields: Fields): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            let processingPromise: Promise<Field> = undefined;

            switch (fieldConfig.FieldTypeKind) {
                case undefined:
                case "":
                case null:
                    Reject(reject, `Field type kind could not be resolved for the field with the internal name ${fieldConfig.InternalName}`, fieldConfig.Title);
                    break;
                case FieldTypeKind.Lookup:
                    processingPromise = this.addLookupField(fieldConfig, parentFields);
                    break;
                case FieldTypeKind.Calculated:
                    processingPromise = this.addCalculatedField(fieldConfig, parentFields);
                    break;
                default:
                    processingPromise = this.addDefaultField(fieldConfig, parentFields);
                    break;
            }

            if (processingPromise) {
                processingPromise
                    .then((fieldProcessingResult) => { resolve(fieldProcessingResult); })
                    .catch((error) => { reject(error); });
            }
        });
    }

    private addDefaultField(fieldConfig: IField, parentFields: Fields) {
        return new Promise<Field>((resolve, reject) => {
            let propertyHash = this.createProperties(fieldConfig);
            parentFields.add(fieldConfig.InternalName, "SP.Field", propertyHash)
                .then((fieldAddResult) => {
                    this.updateField(fieldConfig, fieldAddResult.field)
                        .then((fieldUpdateResult) => { Resolve(resolve, `Added field: '${fieldConfig.InternalName}'`, fieldConfig.Title, fieldUpdateResult); })
                        .catch((error) => { Reject(reject, `Error while adding and updating field with the internal name '${fieldConfig.InternalName}': ` + error, fieldConfig.Title); });
                })
                .catch((error) => { Reject(reject, `Error while adding field with the internal name '${fieldConfig.InternalName}': ` + error, fieldConfig.Title); });
        });
    }

    private addCalculatedField(fieldConfig: IField, parentFields: Fields) {
        return new Promise<Field>((resolve, reject) => {
            let propertyHash = this.createProperties(fieldConfig);
            parentFields.addCalculated(fieldConfig.InternalName, fieldConfig.Formula, Types.DateTimeFieldFormatType[fieldConfig.DateFormat], Types.FieldTypes[fieldConfig.OutputType], propertyHash)
                .then((fieldAddResult) => { Resolve(resolve, `Added field: '${fieldConfig.InternalName}'`, fieldConfig.Title, fieldAddResult.field); })
                .catch((error) => { Reject(reject, `Error while adding field with the internal name '${fieldConfig.InternalName}': ` + error, fieldConfig.Title); });
        });
    }

    private addLookupField(fieldConfig: IField, parentFields: Fields) {
        return new Promise<Field>((resolve, reject) => {
            Reject(reject, "Add lookup field not implemented yet", fieldConfig.Title);
        });
    }

    private updateField(fieldConfig: IField, field: Field): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            let properties = this.createProperties(fieldConfig);
            field.update(properties)
                .then((fieldUpdateResult) => { Resolve(resolve, `Updated field: '${fieldConfig.InternalName}'`, fieldConfig.Title, fieldUpdateResult.field); })
                .catch((error) => { Reject(reject, `Error while updating field with internal name '${fieldConfig.InternalName}': ` + error, fieldConfig.Title); });
        });
    }

    private deleteField(fieldConfig: IField, field: Field): Promise<Field> {
        return new Promise<Field>((resolve, reject) => {
            field.delete()
                .then(() => { Resolve(resolve, `Deleted field: '${fieldConfig.InternalName}'`, fieldConfig.Title); })
                .catch((error) => { Reject(reject, `Error while deleting field with internal name '${fieldConfig.InternalName}': ` + error, fieldConfig.Title); });
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
                delete parsedObject.FieldTypeKind;
                delete parsedObject.DateFormat;
                delete parsedObject.OutputType;
                delete parsedObject.Formula;
                break;
            default:
                delete parsedObject.ControlOption;
                delete parsedObject.Title;
                delete parsedObject.InternalName;
                delete parsedObject.DateFormat;
                delete parsedObject.Formula;
                delete parsedObject.OutputType;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}

