import { Queryable } from "ais-sp-pnp-js/lib/sharepoint/rest/queryable";
import { IPromiseResult } from "../iPromiseResult";

export interface ISPObjectHandler {
    execute(config: any, parent?: Promise<IPromiseResult<void | Queryable>>): Promise<IPromiseResult<void | Queryable>>;
}
