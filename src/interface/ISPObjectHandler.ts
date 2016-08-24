export interface ISPObjectHandler{
    execute(config: any): Promise<any>;
}