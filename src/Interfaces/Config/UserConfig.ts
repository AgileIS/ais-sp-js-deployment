import { AuthenticationType } from "../../Constants/AuthenticationType";

export interface IUserConfig {
    username: string;
    password: string;
    workstation: string;
    authtype: AuthenticationType;
    proxyUrl: string;
}
