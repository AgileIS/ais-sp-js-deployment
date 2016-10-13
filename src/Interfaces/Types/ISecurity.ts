import { IRoleAssignment } from "./iRoleAssignment";

export interface ISecurity {
    BreakRoleInheritance: boolean;
    CopyRoleAssignments: boolean;
    ClearSubscopes: boolean;
    RoleAssignments: Array<IRoleAssignment>;
}
