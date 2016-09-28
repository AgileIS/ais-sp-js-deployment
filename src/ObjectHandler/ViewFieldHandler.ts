import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { View } from "@agileis/sp-pnp-js/lib/sharepoint/rest/views";
import { ISPObjectHandler } from "../Interfaces/ObjectHandler/ISPObjectHandler";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { IViewField } from "../Interfaces/Types/IViewField";
import { Util } from "../Util/Util";

export class ViewFieldHandler implements ISPObjectHandler {
    public execute(viewFieldConfig: IViewField, parentPromise: Promise<IPromiseResult<View>>): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            parentPromise.then((promiseResult) => {
                if (!promiseResult || !promiseResult.value) {
                    Util.Reject<void>(reject, viewFieldConfig.InternalFieldName,
                        `View field handler parent promise value result is null or undefined for the view field with the internal name '${viewFieldConfig.InternalFieldName}'!`);
                } else {
                    let view = promiseResult.value;
                    this.processingViewFieldConfig(viewFieldConfig, view)
                        .then(() => { resolve(); })
                        .catch((error) => {
                            Util.Retry(error, viewFieldConfig.InternalFieldName,
                                () => {
                                    return this.processingViewFieldConfig(viewFieldConfig, view);
                                });
                        });
                }
            });
        });
    }

    private processingViewFieldConfig(viewFieldConfig: IViewField, targetView: View): Promise<IPromiseResult<void>> {
        return new Promise<IPromiseResult<void>>((resolve, reject) => {
            Logger.write(`Processing Add view field: '${viewFieldConfig.InternalFieldName}'.`, Logger.LogLevel.Info);
            targetView.fields.add(viewFieldConfig.InternalFieldName)
                .then(() => { Util.Resolve<void>(resolve, viewFieldConfig.InternalFieldName, `Added view field: '${viewFieldConfig.InternalFieldName}'.`); })
                .catch((error) => {
                    Util.Reject<void>(reject, viewFieldConfig.InternalFieldName,
                        `Error while adding view field with the internal name '${viewFieldConfig.InternalFieldName}': ` + error);
                });
        });
    }
}
