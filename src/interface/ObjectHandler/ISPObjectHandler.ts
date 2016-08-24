export interface ISPObjectHandler{
    execute(config: any, url: string ,parent: Promise<any>): Promise<any>;
}