import { IInstance } from "../Types/IInstance";

export interface ISPObjectHandler{
    execute(config: any, url: string, parent: Promise<IInstance>): Promise<IInstance>;
}