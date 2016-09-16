import { Logger } from "sp-pnp-js/lib/utils/logging";
import { View } from "sp-pnp-js/lib/sharepoint/rest/views";
import { ISPObjectHandler } from "../interface/ObjectHandler/ispobjecthandler";
import { IViewField } from "../interface/Types/IViewField";
import { Reject, Resolve } from "../Util/Util";

export class ViewFieldHandler implements ISPObjectHandler {
    public execute(viewFieldConfig: IViewField, parentPromise: Promise<View>): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            parentPromise.then((parentView) => {
                this.processingViewFieldConfig(viewFieldConfig, parentView)
                    .then(() => { resolve(); })
                    .catch((error) => { reject(error); });
            });
        });
    }

    private processingViewFieldConfig(viewFieldConfig: IViewField, parentView: View): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            Logger.write(`Processing Add view field: '${viewFieldConfig.InternalFieldName}'`, Logger.LogLevel.Info);
            parentView.fields.add(viewFieldConfig.InternalFieldName)
                .then(() => { Resolve(resolve, `Added view field: '${viewFieldConfig.InternalFieldName}'`, viewFieldConfig.InternalFieldName); })
                .catch((error) => { Reject(reject, `Error while adding view field with the internal name '${viewFieldConfig.InternalFieldName}': ` + error, viewFieldConfig.InternalFieldName); });
        });
    }
}
