import {Logger} from "sp-pnp-js/lib/utils/logging";
import {Web} from "sp-pnp-js/lib/sharepoint/rest/webs";

export function ViewFieldRetry(pSpWeb: Web, pListId: string, pParentTitle: string, pElementName: string, pTimeout: number): Promise<void> {
     let promise: Promise<void>;
    setTimeout(() => {
         promise = pSpWeb.lists.getById(pListId).views.getByTitle(pParentTitle).fields.add(pElementName);
    }, pTimeout);
    return promise;
}

export function Resolve(resolve: any, error: string, configElementName: string, value?: any) {
    let errorMsg = `${error} - '${configElementName}'`;
    Logger.write(errorMsg, Logger.LogLevel.Info);
    resolve(value);
}

export function Reject(reject: any, error: string, configElementName: string, value?: any) {
    let errorMsg = `${error} - '${configElementName}'`;
    Logger.write(errorMsg, Logger.LogLevel.Info);
    reject(value);
}