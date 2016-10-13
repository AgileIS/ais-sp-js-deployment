import { AuthenticationType } from "../../constants/authenticationType";

export interface IUserConfig {
    username: string;
    password: string;
    workstation: string;
    authtype: AuthenticationType;
    proxyUrl: string;
}
