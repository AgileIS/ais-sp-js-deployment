import { Logger } from "ais-sp-pnp-js/lib/utils/logging";
import { Web } from "ais-sp-pnp-js/lib/sharepoint/rest/webs";
import { ISPObjectHandler } from "../interfaces/objectHandler/iSpObjectHandler";
import { IFeature } from "../interfaces/types/iFeature";
import { IPromiseResult } from "../interfaces/iPromiseResult";
import { Util } from "../util/util";

export class FeatureHandler implements ISPObjectHandler {
    private noRetry: boolean = false;
    private handlerName = "FeatureHandler";

    public execute(featureConfig: IFeature, parentPromise: Promise<IPromiseResult<Web>>): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            parentPromise.then((promiseResult) => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, this.handlerName,
                        `Feature handler parent promise value result is null or undefined for the feature with the id '${featureConfig.Id}'!`);
                } else {
                    if (featureConfig.Id) {
                        let context = SP.ClientContext.get_current();
                        Util.tryToProcess(featureConfig.Id, () => { return this.processingFeatureConfig(featureConfig, context); }, this.handlerName)
                            .then((featureProcessingResult) => { resolve(featureProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else {
                        Util.Reject<void>(reject, this.handlerName, `Error while processing feature: Feature id is undefined.`);
                    }
                }
            });
        });
    }

    private processingFeatureConfig(featureConfig: IFeature, clientContext: SP.ClientContext): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let processingCrlOption = featureConfig.Deactivate ? "Deactivate" : "Active";
            Logger.write(`${this.handlerName} - Processing ${processingCrlOption} feature: '${featureConfig.Id}'.`, Logger.LogLevel.Info);

            let featureCollection = clientContext.get_site().get_features();
            let currUser = clientContext.get_web().get_currentUser();
            if (SP.FeatureDefinitionScope[featureConfig.Scope.toLocaleLowerCase()] === SP.FeatureDefinitionScope.web) {
                featureCollection = clientContext.get_web().get_features();
            }

            let feature = featureCollection.getById(new SP.Guid(featureConfig.Id));
            clientContext.load(feature);
            clientContext.load(currUser);
            clientContext.executeQueryAsync(
                (sender, args) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void>> = undefined;
                    if (!feature.get_serverObjectIsNull()) {
                        Logger.write(`${this.handlerName} - Found Feature with id: '${featureConfig.Id}'`);
                        if (featureConfig.Deactivate) {
                            processingPromise = this.deactivateFeature(featureConfig, featureCollection, currUser);
                        } else {
                            Util.Resolve<void>(resolve, this.handlerName, `Feature with the id '${featureConfig.Id}' does not have to be added, because it already exists.`);
                            rejectOrResolved = true;
                        }
                    } else {
                        if (featureConfig.Deactivate) {
                            Util.Resolve<void>(resolve, this.handlerName, `Feature with id '${featureConfig.Id}' does not have to be deactivated, because it was not activated.`);
                            rejectOrResolved = true;
                        } else {
                            processingPromise = this.activateFeature(featureConfig, featureCollection, currUser);
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((listProcessingResult) => { resolve(listProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write(`${this.handlerName} - Processing promise is undefined!`, Logger.LogLevel.Error);
                    }
                },
                (sender, args) => {
                    Util.Reject<void>(reject, this.handlerName, `Error while requesting feature with the id '${featureConfig.Id}': `
                            + `${Util.getErrorMessageFromQuery(args.get_message(),args.get_stackTrace())}`);
                }
            );
        });
    }

    private activateFeature(featureConfig: IFeature, featureCollection: SP.FeatureCollection, currentUser: SP.User): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let activateForce = featureConfig.Force === true;
            let scope = featureConfig.Scope ? SP.FeatureDefinitionScope[featureConfig.Scope.toLowerCase()] : SP.FeatureDefinitionScope.none;
            scope = scope === SP.FeatureDefinitionScope.web ? SP.FeatureDefinitionScope.none : scope;
            scope = scope === SP.FeatureDefinitionScope.site ? SP.FeatureDefinitionScope.farm : scope;
            if (((scope === SP.FeatureDefinitionScope.site || scope === SP.FeatureDefinitionScope.farm) && currentUser.get_isSiteAdmin()) || scope === SP.FeatureDefinitionScope.none) {
                featureCollection.add(new SP.Guid(featureConfig.Id), activateForce, scope as SP.FeatureDefinitionScope);
                featureCollection.get_context().executeQueryAsync(
                    (sender, args) => {
                        Util.Resolve<void>(resolve, this.handlerName, `Activated feature: '${featureConfig.Id}'.`);
                    },
                    (sender, args) => {
                        Util.Reject<void>(reject, this.handlerName, `Error while activating feature with the id '${featureConfig.Id}': `
                            + `${Util.getErrorMessageFromQuery(args.get_message(),args.get_stackTrace())}`);
                    });
            } else {
                this.noRetry = true;
                Util.Reject<void>(reject, this.handlerName,
                    `Error while activating feature with the id '${featureConfig.Id}' and feature scope '${featureConfig.Scope}': User '${currentUser.get_loginName()}' is no site administrator`);
            }
        });
    }

    private deactivateFeature(featureConfig: IFeature, featureCollection: SP.FeatureCollection, currentUser: SP.User): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let scope = SP.FeatureDefinitionScope[featureConfig.Scope.toLowerCase()];
            if (((scope === SP.FeatureDefinitionScope.site || scope === SP.FeatureDefinitionScope.farm) && currentUser.get_isSiteAdmin()) || scope === SP.FeatureDefinitionScope.none) {
                featureCollection.remove(new SP.Guid(featureConfig.Id), true);
                featureCollection.get_context().executeQueryAsync(
                    (sender, args) => {
                        Util.Resolve<void>(resolve, this.handlerName, `Deactivated feature: '${featureConfig.Id}'.`);
                    },
                    (sender, args) => {
                        Util.Reject<void>(reject, this.handlerName, `Error while deactivating feature with the id '${featureConfig.Id}': `
                            + `${Util.getErrorMessageFromQuery(args.get_message(),args.get_stackTrace())}`);
                    }
                );
            } else {
                this.noRetry = true;
                Util.Reject<void>(reject, this.handlerName,
                    `Error while deactivating feature with the id '${featureConfig.Id}' and feature scope '${featureConfig.Scope}': User '${currentUser.get_loginName()}' is no site administrator`);
            }
        });
    }
}
