export interface ISPObjectHandler{
    execute(config: any, url: string): Promise<any>;
}