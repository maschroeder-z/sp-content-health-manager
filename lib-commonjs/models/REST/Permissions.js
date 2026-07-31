"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SHAREPOINT_RELEVANT_ENTRA_ROLES = exports.PermissionKind = exports.PrincipalType = exports.SharePointArtefactType = void 0;
var SharePointArtefactType;
(function (SharePointArtefactType) {
    SharePointArtefactType["Web"] = "Web";
    SharePointArtefactType["List"] = "List";
    SharePointArtefactType["ListItem"] = "ListItem";
})(SharePointArtefactType || (exports.SharePointArtefactType = SharePointArtefactType = {}));
var PrincipalType;
(function (PrincipalType) {
    PrincipalType[PrincipalType["None"] = 0] = "None";
    PrincipalType[PrincipalType["User"] = 1] = "User";
    PrincipalType[PrincipalType["DistributionList"] = 2] = "DistributionList";
    PrincipalType[PrincipalType["SecurityGroup"] = 4] = "SecurityGroup";
    PrincipalType[PrincipalType["SharePointGroup"] = 8] = "SharePointGroup";
    PrincipalType[PrincipalType["All"] = 15] = "All";
})(PrincipalType || (exports.PrincipalType = PrincipalType = {}));
// Curated subset of SP.PermissionKind (real CSOM numeric values, 1-based bit positions).
var PermissionKind;
(function (PermissionKind) {
    PermissionKind[PermissionKind["ViewListItems"] = 1] = "ViewListItems";
    PermissionKind[PermissionKind["AddListItems"] = 2] = "AddListItems";
    PermissionKind[PermissionKind["EditListItems"] = 3] = "EditListItems";
    PermissionKind[PermissionKind["DeleteListItems"] = 4] = "DeleteListItems";
    PermissionKind[PermissionKind["ManageLists"] = 12] = "ManageLists";
    PermissionKind[PermissionKind["ViewPages"] = 18] = "ViewPages";
    PermissionKind[PermissionKind["ManagePermissions"] = 26] = "ManagePermissions";
    PermissionKind[PermissionKind["ManageWeb"] = 31] = "ManageWeb";
    PermissionKind[PermissionKind["FullMask"] = 65] = "FullMask";
})(PermissionKind || (exports.PermissionKind = PermissionKind = {}));
/**
 * Built-in Entra directory roles relevant to a SharePoint permissions audit, with their universal
 * role template GUIDs (identical in every tenant). Verified against Microsoft's built-in roles
 * reference: https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference
 * Querying a role's members requires the RoleManagement.Read.Directory Graph permission, requested
 * in config/package-solution.json and subject to tenant admin approval.
 */
exports.SHAREPOINT_RELEVANT_ENTRA_ROLES = [
    { displayName: 'Global Administrator', roleTemplateId: '62e90394-69f5-4237-9190-012177145e10' },
    { displayName: 'SharePoint Administrator', roleTemplateId: 'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' },
    { displayName: 'Global Reader', roleTemplateId: 'f2ef992c-3afb-46b9-b7cf-a126ee74c451' },
    { displayName: 'User Administrator', roleTemplateId: 'fe930be7-5e62-47db-91af-98c3a49a38b1' },
    { displayName: 'Security Administrator', roleTemplateId: '194ae4cb-b126-40b2-bd5b-6091b380977d' },
    { displayName: 'Teams Administrator', roleTemplateId: '69091246-20e8-4a56-aa4d-066075b2a7a8' },
    { displayName: 'Exchange Administrator', roleTemplateId: '29232cdf-9323-42fd-ade2-1d097af3e4de' },
    { displayName: 'Compliance Administrator', roleTemplateId: '17315797-102d-40b4-93e0-432062caca18' }
];
//# sourceMappingURL=Permissions.js.map