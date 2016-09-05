import {Logger} from "sp-pnp-js/lib/utils/logging";

export function RejectAndLog(pError: string, pElementName: string, reject: any) {
    let errorMsg = `${pError}  - '${pElementName}'`;
    Logger.write(errorMsg, 1);
    reject(errorMsg);
}