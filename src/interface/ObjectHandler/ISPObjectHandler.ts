export interface ISPObjectHandler{
    execute(config: any, url: string, parentConfig: any): Promise<any>;
}