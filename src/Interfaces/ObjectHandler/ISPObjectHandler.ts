import { Queryable } from "@agileis/sp-pnp-js/lib/sharepoint/rest/queryable";
import { IPromiseResult } from "../IPromiseResult";

export interface ISPObjectHandler {
    execute(config: any, parent?: Promise<IPromiseResult<void | Queryable>>): Promise<IPromiseResult<void | Queryable>>;
}