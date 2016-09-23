import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IFeature } from "../interface/Types/IFeature";
import { ControlOption } from "../Constants/ControlOption";
import { Resolve, Reject } from "../Util/Util";

export class FeatureHandler implements ISPObjectHandler {
    public execute(featureConfig: IFeature, parentPromise: Promise<Web>): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            parentPromise.then((parentWeb) => {
                this.processingFeatureConfig(featureConfig, parentWeb)
                    .then((featureProsssingResult) => { resolve(featureProsssingResult); })
                    .catch((error) => { reject(error); });
            });
        });
    }

    private processingFeatureConfig(featureConfig: IFeature, parentWeb: Web): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            let processingCrlOption = featureConfig.ControlOption === ControlOption.Delete
                ? "Deactivate" : featureConfig.ControlOption === ControlOption.Update
                    ? featureConfig.ControlOption : "Active";
            Logger.write(`Processing ${processingCrlOption} feature: '${featureConfig.ID}'`, Logger.LogLevel.Info);

            let context = new SP.ClientContext(parentWeb.toUrl().split("/_")[0]);
            let featureCollection = context.get_site().get_features();
            if (featureConfig.Scope === SP.FeatureDefinitionScope.web) {
                featureCollection = context.get_web().get_features();
            }
            let feature = featureCollection.getById(new SP.Guid(featureConfig.ID));
            context.load(feature);
            context.executeQueryAsync(
                (sender, args) => {
                    let processingPromise: Promise<void> = undefined;
                    if (!feature.get_serverObjectIsNull()) {
                        switch (featureConfig.ControlOption) {
                            case ControlOption.Update:
                                processingPromise = this.updateFeature(featureConfig, featureCollection);
                                break;
                            case ControlOption.Delete:
                                processingPromise = this.deleteFeature(featureConfig, featureCollection);
                                break;
                            default:
                                Resolve(resolve, `Feature with the id '${featureConfig.ID}' already exists`, featureConfig.ID);
                                break;
                        }
                    } else {
                        switch (featureConfig.ControlOption) {
                            case ControlOption.Update:
                            case ControlOption.Delete:
                                Reject(reject, `Feature with id '${featureConfig.ID}' does not exists`, featureConfig.ID);
                                break;
                            default:
                                processingPromise = this.addFeature(featureConfig, featureCollection);
                                break;
                        }
                    }

                    if (processingPromise) {
                        processingPromise
                            .then((listProcessingResult) => { resolve(listProcessingResult); })
                            .catch((error) => { reject(error); });
                    }
                },
                (sender, args) => {
                    Reject(reject, `Error while requesting feature with the id '${featureConfig.ID}': ${args.get_message()} '\n' ${args.get_stackTrace()}`, featureConfig.ID);
                }
            );
        });
    }

    private addFeature(featureConfig: IFeature, featureCollection: SP.FeatureCollection): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            let scope = featureConfig.Scope ? SP.FeatureDefinitionScope[featureConfig.Scope] : SP.FeatureDefinitionScope.none;
            scope = scope === SP.FeatureDefinitionScope.web ? SP.FeatureDefinitionScope.none : scope;
            scope = scope === SP.FeatureDefinitionScope.site ? SP.FeatureDefinitionScope.farm : scope;
            featureCollection.add(new SP.Guid(featureConfig.ID), true, scope as SP.FeatureDefinitionScope);
            featureCollection.get_context().executeQueryAsync(
                (sender, args) => {
                    Resolve(resolve, `Activated feature: '${featureConfig.ID}'`, featureConfig.ID);
                },
                (sender, args) => {
                    Reject(reject, `Error while activating feature with the id '${featureConfig.ID}': ${args.get_message()} '\n' ${args.get_stackTrace()}`, featureConfig.ID);
                });
        });
    }

    private updateFeature(featureConfig: IFeature, featureCollection: SP.FeatureCollection): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Reject(reject, "Updating feature is not possible.", featureConfig.ID);
        });
    }

    private deleteFeature(featureConfig: IFeature, featureCollection: SP.FeatureCollection): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            featureCollection.remove(new SP.Guid(featureConfig.ID), true);
            featureCollection.get_context().executeQueryAsync(
                (sender, args) => {
                    Resolve(resolve, `Deactivated feature: '${featureConfig.ID}'`, featureConfig.ID);
                },
                (sender, args) => {
                    Reject(reject, `Error while deactivating feature with the id '${featureConfig.ID}': ${args.get_message()} '\n' ${args.get_stackTrace()}`, featureConfig.ID);
                }
            );
        });
    }
}
