import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";

export function ViewFieldRetry(pSpWeb: Web, pListId: string, pParentTitle: string, pElementName: string, pTimeout: number): Promise<void> {
     let promise: Promise<void>;
    setTimeout(() => {
         promise = pSpWeb.lists.getById(pListId).views.getByTitle(pParentTitle).fields.add(pElementName);
    }, pTimeout);
    return promise;
}

export function Resolve(resolve: any, msg: string, configElementName: string, value?: any) {
    let errorMsg = `'${configElementName}' - ${msg}`;
    Logger.write(errorMsg, Logger.LogLevel.Info);

    let resolveValue = msg;
    if (value) {
        resolveValue = value;
    }
    resolve(resolveValue);
}

export function Reject(reject: any, msg: string, configElementName: string, value?: any) {
    let errorMsg = `'${configElementName}' - ${msg}`;
    Logger.write(errorMsg, Logger.LogLevel.Info);

    let rejectValue = msg;
    if (value) {
        rejectValue = value;
    }
    reject(rejectValue);
}
