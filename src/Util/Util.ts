import { Logger } from "@agileis/sp-pnp-js/lib/utils/logging";
import * as url from "url";
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
            Logger.write(errorMsg, Logger.LogLevel.Error);
        }

        reject(new PromiseResult<T>(promiseResultMessage, promiseResultValue));
    }

    export function JoinAndNormalizeUrl(urlParts: Array<string>): string {
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

    export function getErrorMessage(error: any): any {
        let errorMessage = error;
        if (typeof error === "object") {
            if ((error as Object).hasOwnProperty("message")) { errorMessage = error.message; }
        }
        return errorMessage;
    }

    export function Retry(error: any, configNodeIdentifier: string, retryFunction: () => Promise<IPromiseResult<any>>) {
        Logger.write(`Retry process for '${configNodeIdentifier}' because Error: ${getErrorMessage(error)}`, Logger.LogLevel.Warning);
        setTimeout(() => {
            Logger.write(`Retry first time: '${configNodeIdentifier}'`, Logger.LogLevel.Warning);
            retryFunction().then((result) => {
                return Promise.resolve(result);
            }).catch((firstRetryError) => {
                setTimeout(() => {
                    Logger.write(`Retry failed first time for '${configNodeIdentifier}' - ${getErrorMessage(firstRetryError)}`, Logger.LogLevel.Warning);
                    retryFunction().then((result) => {
                        return Promise.resolve(result);
                    }).catch((secondRetryError) => {
                        Logger.write(`Retry failed second time: '${configNodeIdentifier}' - Reject`, Logger.LogLevel.Warning);
                        return Promise.reject(getErrorMessage(secondRetryError));
                    });
                }, 5000);
            });
        }, 2500);
    }

    export function getRelativeUrl(absoluteUrl: string): string {
        let urlObject = url.parse(absoluteUrl, true, true);

        let relativeUrl = urlObject.pathname;

        if (urlObject.search) {
            relativeUrl += urlObject.search;
        }

        if (urlObject.hash) {
            relativeUrl += urlObject.hash;
        }

        return relativeUrl;
    }

    export function trimEnd(content: string, trimEndChar: string): string {
        if (trimEndChar.length > 1) {
            throw new Error("Argument 'trimEndChar' value is invalid. Length is greater then one!");
        }

        return content.charAt(content.length - 1) === trimEndChar ? content.substring(0, content.length - 1) : content;
    }

    export function trimStart(content: string, trimStarChar: string): string {
        if (trimStarChar.length > 1) {
            throw new Error("Argument 'trimStarChar' value is invalid. Length is greater then one!");
        }

        return content.charAt(0) === trimStarChar ? content.substring(1, content.length) : content;
    }

    function replaceSiteToken(content: string, siteRelativeUrl: string): string {
        return content.replace(/~replaceSite/g, siteRelativeUrl);
    }

    function replaceLayoutsToken(content: string, siteRelativeUrlWithLayouts: string): string {
        return content.replace(/~replaceLayouts/g, siteRelativeUrlWithLayouts);
    }

    function replaceEncodeSiteToken(content: string, siteRelativeUrl: string): string {
        return content.replace(/~replaceEncodeSite/g, encodeURIComponent(siteRelativeUrl));
    }

    function replaceEncodedLayoutsToken(content: string, siteRelativeUrlWithLayouts: string): string {
        return content.replace(/~replaceEncodeLayouts/g, siteRelativeUrlWithLayouts);
    }

    export function replaceUrlTokens(content: string, siteRelativeUrl: string, layoutsUrlPart: string): string {
        if (!layoutsUrlPart) {
            throw new Error("Argument 'layoutsUrlPart' value is undefined");
        }

        if (!siteRelativeUrl) {
            throw new Error("Argument 'siteRelativeUrl' value is undefined");
        }

        let relativeUrl = trimEnd(siteRelativeUrl, "/");
        let relativeUrlWithLayouts = `${relativeUrl}/${trimEnd(trimStart(layoutsUrlPart, "/"), "/")}`;

        let outputContent = replaceSiteToken(content, relativeUrl);
        outputContent = replaceLayoutsToken(outputContent, relativeUrlWithLayouts);

        outputContent = replaceEncodeSiteToken(outputContent, relativeUrl);
        outputContent = replaceEncodedLayoutsToken(outputContent, relativeUrlWithLayouts);

        return outputContent;
    }
}
