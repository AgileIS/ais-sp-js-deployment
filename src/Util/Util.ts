import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { Web } from "@agileis/sp-pnp-js/lib/sharepoint/rest/webs";
import { PromiseResult } from "../PromiseResult";
import { Queryable } from "@agileis/sp-pnp-js/lib/sharepoint/rest/queryable";

export namespace Util {
    export function ViewFieldRetry(pSpWeb: Web, pListId: string, pParentTitle: string, pElementName: string, pTimeout: number): Promise<void> {
        let promise: Promise<void>;
        setTimeout(() => {
            promise = pSpWeb.lists.getById(pListId).views.getByTitle(pParentTitle).fields.add(pElementName);
        }, pTimeout);
        return promise;
    }

    /** Resolve with a IPromiseResult */
    export function Resolve<T>(resolve: (value?: PromiseResult<T> | Thenable<PromiseResult<T>>) => void, configNodeIdentifier: string, promiseResultMessage: string, promiseResultValue?: T) {
        if (configNodeIdentifier && promiseResultMessage) {
            let errorMsg = `'${configNodeIdentifier}' - ${promiseResultMessage}`;
            Logger.write(errorMsg, Logger.LogLevel.Info);
        }

        resolve(new PromiseResult<T>(promiseResultMessage, promiseResultValue));
    }

    /** Reject with a IPromiseResult */
    export function Reject<T>(reject: (error?: any) => void, configNodeIdentifier: string, promiseResultMessage: string, promiseResultValue?: T) {
        if (configNodeIdentifier && promiseResultMessage) {
            let errorMsg = `'${configNodeIdentifier}' - ${promiseResultMessage}`;
            Logger.write(errorMsg, Logger.LogLevel.Info);
        }

        reject(new PromiseResult<T>(promiseResultMessage, promiseResultValue));
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
}
