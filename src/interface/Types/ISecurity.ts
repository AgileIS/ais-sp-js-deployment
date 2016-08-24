import { IRoleAssignment } from "./IRoleAssignment";

export interface ISecurity {
    BreakRoleInheritance: boolean;
    CopyRoleAssignments: boolean;
    ClearSubscopes: boolean;
    RoleAssignments: Array<IRoleAssignment>;
}