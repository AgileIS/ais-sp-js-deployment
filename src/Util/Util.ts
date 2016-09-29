import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import { IPromiseResult } from "../Interfaces/IPromiseResult";
import { PromiseResult } from "../PromiseResult";

export namespace Util {
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

    function getErrorMessage(error: any): any{
        let errorMessage = error;
        if (typeof error === "object") { if ((error as Object).hasOwnProperty("message")) { errorMessage = error.message; }
        }
        return errorMessage;
    }

    export function Retry(error: any, configNodeIdentifier: string, retryFunction: () => Promise<IPromiseResult<any>>) {
        Logger.write(`Retry process for '${configNodeIdentifier}' because Error: ${getErrorMessage(error)}`);
        setTimeout(() => {
            Logger.write(`Retry first time: '${configNodeIdentifier}'`);
            retryFunction().then((result) => {
                return Promise.resolve(result);
            }).catch((firstRetryError) => {
                setTimeout(() => {
                    Logger.write(`Retry failed first time for '${configNodeIdentifier}' - ${getErrorMessage(firstRetryError)}`);
                    retryFunction().then((result) => {
                        return Promise.resolve(result);
                    }).catch((secondRetryError) => {
                        Logger.write(`Retry failed second time: '${configNodeIdentifier}' - Reject`);
                        return Promise.reject(getErrorMessage(secondRetryError));
                    });
                }, 5000);
            });
        }, 2500);
    }
}
