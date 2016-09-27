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

export function UrlJoin(urlParts: Array<string>): string {
    let normalizedUrl = urlParts.join("/");
    let parts = normalizedUrl.split("/");
    parts[0] = parts[0].concat(":").replace("::", ":");
    normalizedUrl = parts.join("/").replace("//", "/");
    normalizedUrl = normalizedUrl.replace(/:\//g, "://");
    normalizedUrl = normalizedUrl.replace(/\/$/, "");
    // normalizedUrl = normalizedUrl.replace(/\/(\?|&|#[^!])/g, "$1");
    // normalizedUrl = normalizedUrl.replace(/(\?.+)\?/g, "$1&");

    return normalizedUrl;
}
