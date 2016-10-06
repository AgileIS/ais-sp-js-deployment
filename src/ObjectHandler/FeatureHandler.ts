import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IFeature } from "../Interfaces/Types/IFeature";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { Util } from "../Util/Util";

export class FeatureHandler implements ISPObjectHandler {
    public execute(featureConfig: IFeature, parentPromise: Promise<IPromiseResult<Web>>): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            parentPromise.then((promiseResult) => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, featureConfig.Id,
                        `Feature handler parent promise value result is null or undefined for the feature with the id '${featureConfig.Id}'!`);
                } else {
                    if (featureConfig.Id) {
                        let context = SP.ClientContext.get_current();
                        this.processingFeatureConfig(featureConfig, context)
                            .then((featureProsssingResult) => { resolve(featureProsssingResult); })
                            .catch((error) => {
                                Util.Retry(error, featureConfig.Id,
                                    () => {
                                        return this.processingFeatureConfig(featureConfig, context);
                                    });
                            });
                    } else {
                        Util.Reject<void>(reject, "Unknow feature", `Error while processing feature: Feature id is undefined.`);
                    }
                }
            });
        });
    }

    private processingFeatureConfig(featureConfig: IFeature, clientContext: SP.ClientContext): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let processingCrlOption = featureConfig.Deactivate ? "Deactivate" : "Active";
            Logger.write(`Processing ${processingCrlOption} feature: '${featureConfig.Id}'.`, Logger.LogLevel.Info);

            let featureCollection = clientContext.get_site().get_features();
            if (featureConfig.Scope === SP.FeatureDefinitionScope.web.toString()) {
                featureCollection = clientContext.get_web().get_features();
            }

            let feature = featureCollection.getById(new SP.Guid(featureConfig.Id));
            clientContext.load(feature);
            clientContext.executeQueryAsync(
                (sender, args) => {
                    let rejectOrResolved = false;
                    let processingPromise: Promise<IPromiseResult<void>> = undefined;
                    if (!feature.get_serverObjectIsNull()) {
                        Logger.write(`Found Feature with id: '${featureConfig.Id}'`);
                        if (featureConfig.Deactivate) {
                            processingPromise = this.deactivateFeature(featureConfig, featureCollection);
                        } else {
                            Util.Resolve<void>(resolve, featureConfig.Id, `Feature with the id '${featureConfig.Id}' does not have to be added, because it already exists.`);
                            rejectOrResolved = true;
                        }
                    } else {
                        if (featureConfig.Deactivate) {
                            Util.Resolve<void>(resolve, featureConfig.Id, `Feature with id '${featureConfig.Id}' does not have to be deactivated, because it was not activated.`);
                            rejectOrResolved = true;
                        } else {
                            processingPromise = this.activateFeature(featureConfig, featureCollection);
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((listProcessingResult) => { resolve(listProcessingResult); })
                            .catch((error) => { reject(error); });
                    } else if (!rejectOrResolved) {
                        Logger.write("Feature handler processing promise is undefined!", Logger.LogLevel.Error);
                    }
                },
                (sender, args) => {
                    Util.Reject<void>(reject, featureConfig.Id, `Error while requesting feature with the id '${featureConfig.Id}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                }
            );
        });
    }

    private activateFeature(featureConfig: IFeature, featureCollection: SP.FeatureCollection): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            let scope = featureConfig.Scope ? SP.FeatureDefinitionScope[featureConfig.Scope.toLowerCase()] : SP.FeatureDefinitionScope.none;
            scope = scope === SP.FeatureDefinitionScope.web ? SP.FeatureDefinitionScope.none : scope;
            scope = scope === SP.FeatureDefinitionScope.site ? SP.FeatureDefinitionScope.farm : scope;
            featureCollection.add(new SP.Guid(featureConfig.Id), false, scope as SP.FeatureDefinitionScope);
            featureCollection.get_context().executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<void>(resolve, featureConfig.Id, `Activated feature: '${featureConfig.Id}'.`);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, featureConfig.Id, `Error while activating feature with the id '${featureConfig.Id}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                });
        });
    }

    private deactivateFeature(featureConfig: IFeature, featureCollection: SP.FeatureCollection): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            featureCollection.remove(new SP.Guid(featureConfig.Id), true);
            featureCollection.get_context().executeQueryAsync(
                (sender, args) => {
                    Util.Resolve<void>(resolve, featureConfig.Id, `Deactivated feature: '${featureConfig.Id}'.`);
                },
                (sender, args) => {
                    Util.Reject<void>(reject, featureConfig.Id, `Error while deactivating feature with the id '${featureConfig.Id}': ${args.get_message()} '\n' ${args.get_stackTrace()}`);
                }
            );
        });
    }
}
