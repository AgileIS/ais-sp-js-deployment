import { Queryable } from "sp-pnp-js/lib/sharepoint/rest/queryable";

export interface ISPObjectHandler{
    execute(config: any, parent?: Promise<Queryable>): Promise<Queryable | void>;
}