import {Logger} from "sp-pnp-js/lib/utils/logging";
import * as web from "sp-pnp-js/lib/sharepoint/rest/webs";

export function RejectAndLog(pError: string, pElementName: string, reject: any) {
    let errorMsg = `${pError}  - '${pElementName}'`;
    Logger.write(errorMsg, 1);
    reject(errorMsg);
}

export function ViewFieldRetry(pSpWeb: web.Web, pListId: string, pParentTitle: string, pElementName: string, pTimeout: number): Promise<void> {
     let promise: Promise<void>;
    setTimeout(() => {
         promise = pSpWeb.lists.getById(pListId).views.getByTitle(pParentTitle).fields.add(pElementName);
    }, pTimeout);
    return promise;
}