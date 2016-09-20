import { AuthenticationType } from "../../Constants/AuthenticationType";

export interface UserConfig {
    username: string;
    password: string;
    workstation: string;
    authtype: AuthenticationType;
    proxyUrl: string;
}