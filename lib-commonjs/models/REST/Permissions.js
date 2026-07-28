"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PermissionKind = exports.PrincipalType = exports.SharePointArtefactType = void 0;
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
//# sourceMappingURL=Permissions.js.map