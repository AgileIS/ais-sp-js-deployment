import { IPromiseResult } from "./interfaces/iPromiseResult";

export class PromiseResult<TValue> implements IPromiseResult<TValue> {
    public message: string;
    public value: TValue;
    constructor(message: string, value: TValue) {
        this.message = message;
        this.value = value;
    }
}
