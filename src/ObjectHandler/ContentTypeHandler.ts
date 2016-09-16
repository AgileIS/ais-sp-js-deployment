import { Logger } from "sp-pnp-js/lib/utils/logging";
import { Web } from "sp-pnp-js/lib/sharepoint/rest/webs";
import { ContentType } from "sp-pnp-js/lib/sharepoint/rest/ContentTypes";
import { Util } from "sp-pnp-js/lib/utils/util";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IContentType } from "../interface/Types/IContentType";
import { ControlOption } from "../Constants/ControlOption";
import { Reject, Resolve } from "../Util/Util";

export class ContentTypeHandler implements ISPObjectHandler {
    public execute(contentTypeConfig: IContentType, parentPromise: Promise<Web>): Promise<ContentType> {
        return new Promise<ContentType>((resolve, reject) => {
            parentPromise.then((parentList) => {
                this.processingContentTypeConfig(contentTypeConfig, parentList)
                    .then((contentTypeProsssingResult) => { resolve(contentTypeProsssingResult); })
                    .catch((error) => { reject(error); });
            });
        });
    }

    private processingContentTypeConfig(contentTypeConfig: IContentType, parentWeb: Web): Promise<ContentType> {
        return new Promise<ContentType>((resolve, reject) => {
            Logger.write(`Processing ${contentTypeConfig.ControlOption === ControlOption.Add || contentTypeConfig.ControlOption === undefined ? "Add"
                : contentTypeConfig.ControlOption} content type: '${contentTypeConfig.Name}'`, Logger.LogLevel.Info);
            parentWeb.contentTypes.filter(`Name eq '${contentTypeConfig.Name}'`).get().then((contentTypeRequestResults) => {
                let processingPromise: Promise<ContentType> = undefined;

                if (contentTypeRequestResults && contentTypeRequestResults.length === 1) {
                    let contentType = parentWeb.contentTypes.getById(contentTypeRequestResults[0].Id.StringValue);
                    switch (contentTypeConfig.ControlOption) {
                        case ControlOption.Update:
                            processingPromise = this.updateContentType(contentTypeConfig, contentType);
                            break;
                        case ControlOption.Delete:
                            processingPromise = this.deleteContentType(contentTypeConfig, contentType);
                            break;
                        default:
                            Resolve(resolve, `Content type with the name '${contentTypeConfig.Name}' already exists`, contentTypeConfig.Name, contentType);
                            break;
                    }
                } else {
                    switch (contentTypeConfig.ControlOption) {
                        case ControlOption.Update:
                        case ControlOption.Delete:
                            Reject(reject, `Content type with the name '${contentTypeConfig.Name}' does not exists`, contentTypeConfig.Name);
                            break;
                        default:
                            processingPromise = this.addContentType(contentTypeConfig, parentWeb);
                            break;
                    }
                }

                if (processingPromise) {
                    processingPromise.then((contentTypeProsssingResult) => { resolve(contentTypeProsssingResult); }).catch((error) => { reject(error); });
                }
            }).catch((error) => {
                Reject(reject, `Error while requesting content type with the name '${contentTypeConfig.Name}': ` + error, contentTypeConfig.Name);
            });
        });
    }

    private addContentTypeToCollection(properties: IContentType, parentWeb: Web): Promise<ContentTypeAddResult> {
        let postBody = JSON.stringify(Util.extend({
            "__metadata": { "type": "SP.ContentType" },
        }, properties));

        return parentWeb.contentTypes.post({ body: postBody })
            .then((data) => {
                return { contentType: parentWeb.contentTypes.getById(data.Id), data: data };
            });
    }

    private addContentType(contentTypeConfig: IContentType, parentWeb: Web): Promise<ContentType> {
        return new Promise<ContentType>((resolve, reject) => {
            let properties = this.createProperties(contentTypeConfig);
            this.addContentTypeToCollection(properties, parentWeb)
                .then((contentTypeAddResult) => { Resolve(resolve, `Added content Type: '${contentTypeConfig.Name}'`, contentTypeConfig.Name, contentTypeAddResult.contentType); })
                .catch((error) => { Reject(reject, `Error while adding content type with the name '${contentTypeConfig.Name}': ` + error, contentTypeConfig.Name); });
        });
    }

    private mergeContentType(properties: any, contentType: ContentType): Promise<ContentTypeUpdateResult> {
        let postBody: string = JSON.stringify(Util.extend({
            "__metadata": { "type": "SP.ContentType" },
        }, properties));

        return contentType.post({ body: postBody, headers: { "X-HTTP-Method": "MERGE" } })
            .then((data) => {
                return {
                    contentType: contentType,
                    data: data,
                };
            });
    }

    private updateContentType(contentTypeConfig: IContentType, contentType: ContentType): Promise<ContentType> {
        return new Promise<ContentType>((resolve, reject) => {
            let properties = this.createProperties(contentTypeConfig);
            this.mergeContentType(properties, contentType)
                .then((contentTypeUpdateResult) => { Resolve(resolve, `Updated content type: '${contentTypeConfig.Name}'`, contentTypeConfig.Name, contentTypeUpdateResult.contentType); })
                .catch((error) => { Reject(reject, `Error while updating content type with the name '${contentTypeConfig.Name}': ` + error, contentTypeConfig.Name); });
        });
    }

    private deleteContentType(contentTypeConfig: IContentType, contentType: ContentType): Promise<ContentType> {
        return new Promise<ContentType>((resolve, reject) => {
            contentType.post({ headers: { "X-HTTP-Method": "DELETE" } })
                .then(() => { Resolve(resolve, `Deleted content Type: '${contentTypeConfig.Name}'`, contentTypeConfig.Name); })
                .catch((error) => { Reject(reject, `Error while deleting content type with the name '${contentTypeConfig.Name}': ` + error, contentTypeConfig.Name); });
        });
    }

    private createProperties(contentTypeConfig: IContentType) {
        let stringifiedObject: string;
        stringifiedObject = JSON.stringify(contentTypeConfig);
        let parsedObject: IContentType = JSON.parse(stringifiedObject);
        switch (contentTypeConfig.ControlOption) {
            case ControlOption.Update:
                delete parsedObject.ControlOption;
                break;
            default:
                delete parsedObject.ControlOption;
                break;
        }
        stringifiedObject = JSON.stringify(parsedObject);
        return JSON.parse(stringifiedObject);
    }
}

export interface ContentTypeAddResult {
    contentType: ContentType;
    data: any;
}

export interface ContentTypeUpdateResult {
    contentType: ContentType;
    data: any;
}
