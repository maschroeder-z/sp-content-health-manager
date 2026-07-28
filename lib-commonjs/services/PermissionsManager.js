"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PermissionsManager = void 0;
var tslib_1 = require("tslib");
var sp_http_1 = require("@microsoft/sp-http");
var Permissions_1 = require("../models/REST/Permissions");
var PermissionsManager = /** @class */ (function () {
    function PermissionsManager(spHttpClient) {
        this.spHttpClient = spHttpClient;
    }
    PermissionsManager.prototype.get4ArtefactPermissions = function (artefact) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var base, response, data, roleAssignments, error_1;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        return [4 /*yield*/, this.resolveNearestSecurableBaseUrl(artefact)];
                    case 1:
                        base = _a.sent();
                        return [4 /*yield*/, this.spHttpClient.get("".concat(base, "/roleassignments?$expand=Member,RoleDefinitionBindings"), sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        roleAssignments = this.unwrapCollection(data);
                        return [2 /*return*/, roleAssignments.map(function (roleAssignment) {
                                var _a;
                                var member = roleAssignment.Member || {};
                                var roles = ((_a = roleAssignment.RoleDefinitionBindings) === null || _a === void 0 ? void 0 : _a.results) || [];
                                var principalType = member.PrincipalType;
                                return {
                                    principalId: member.Id,
                                    principalType: principalType,
                                    isGroup: principalType !== Permissions_1.PrincipalType.User,
                                    displayName: member.Title,
                                    loginName: member.LoginName,
                                    email: member.Email || undefined,
                                    roles: roles.map(function (role) { return role.Name; })
                                };
                            })];
                    case 4:
                        error_1 = _a.sent();
                        console.error('Error retrieving artefact permissions:', error_1);
                        throw error_1;
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    PermissionsManager.prototype.checkPermission4User = function (user, artefact) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var loginHint, loginName, base, encodedLoginName, response, data, entity, mask, error_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        loginHint = user.mail || user.userPrincipalName;
                        if (!loginHint) {
                            throw new Error('User must have a mail or userPrincipalName to resolve a SharePoint login name.');
                        }
                        return [4 /*yield*/, this.resolveLoginName(artefact.webUrl, loginHint)];
                    case 1:
                        loginName = _a.sent();
                        base = this.buildBaseUrl(artefact);
                        encodedLoginName = encodeURIComponent("'".concat(loginName.replace(/'/g, "''"), "'"));
                        return [4 /*yield*/, this.spHttpClient.get("".concat(base, "/getusereffectivepermissions(@v)?@v=").concat(encodedLoginName), sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        entity = this.unwrapEntity(data);
                        mask = {
                            high: Number((entity === null || entity === void 0 ? void 0 : entity.High) || 0),
                            low: Number((entity === null || entity === void 0 ? void 0 : entity.Low) || 0)
                        };
                        return [2 /*return*/, {
                                loginName: loginName,
                                canView: this.hasPermissionKind(mask, Permissions_1.PermissionKind.ViewListItems),
                                canContribute: this.hasPermissionKind(mask, Permissions_1.PermissionKind.AddListItems) && this.hasPermissionKind(mask, Permissions_1.PermissionKind.EditListItems),
                                canEdit: this.hasPermissionKind(mask, Permissions_1.PermissionKind.EditListItems),
                                canDelete: this.hasPermissionKind(mask, Permissions_1.PermissionKind.DeleteListItems),
                                canManageLists: this.hasPermissionKind(mask, Permissions_1.PermissionKind.ManageLists),
                                canManagePermissions: this.hasPermissionKind(mask, Permissions_1.PermissionKind.ManagePermissions),
                                hasFullControl: this.hasPermissionKind(mask, Permissions_1.PermissionKind.FullMask),
                                rawMask: mask
                            }];
                    case 4:
                        error_2 = _a.sent();
                        console.error('Error checking user permission:', error_2);
                        throw error_2;
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Resolves the SharePointArtefact (list + item) that owns the file/page at the given URL.
     * Needed because identifiers exposed by other APIs (e.g. Microsoft Graph's sitePage id) do not
     * necessarily match the underlying SharePoint list item id.
     */
    PermissionsManager.prototype.resolveArtefactFromFileUrl = function (webUrl, fileUrl) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var normalizedWebUrl, serverRelativeUrl, encodedServerRelativeUrl, response, data, entity, listId, itemId, error_3;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        normalizedWebUrl = this.normalizeWebUrl(webUrl);
                        serverRelativeUrl = this.toServerRelativeUrl(fileUrl);
                        encodedServerRelativeUrl = serverRelativeUrl
                            .replace(/'/g, "''")
                            .replace(/\(/g, '%28')
                            .replace(/\)/g, '%29');
                        return [4 /*yield*/, this.spHttpClient.get("".concat(normalizedWebUrl, "/_api/web/GetFileByServerRelativeUrl('").concat(encodedServerRelativeUrl, "')/ListItemAllFields?$select=Id,ParentList/Id&$expand=ParentList"), sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 1:
                        response = _b.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _b.sent();
                        entity = this.unwrapEntity(data);
                        listId = (_a = entity === null || entity === void 0 ? void 0 : entity.ParentList) === null || _a === void 0 ? void 0 : _a.Id;
                        itemId = Number(entity === null || entity === void 0 ? void 0 : entity.Id);
                        if (!listId || Number.isNaN(itemId)) {
                            throw new Error("Could not resolve list item for file URL: ".concat(fileUrl));
                        }
                        return [2 /*return*/, {
                                type: Permissions_1.SharePointArtefactType.ListItem,
                                webUrl: normalizedWebUrl,
                                listId: listId,
                                itemId: itemId
                            }];
                    case 3:
                        error_3 = _b.sent();
                        console.error('Error resolving artefact from file URL:', error_3);
                        throw error_3;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    PermissionsManager.prototype.hasUniquePermission = function (artefact) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var error_4;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getHasUniqueRoleAssignments(this.buildBaseUrl(artefact))];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2:
                        error_4 = _a.sent();
                        console.error('Error checking unique permissions:', error_4);
                        throw error_4;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Walks up Web -> List -> ListItem until it finds the object that actually owns the permissions
     * (HasUniqueRoleAssignments === true), since an inheriting object's own `roleassignments` collection
     * is empty rather than mirroring its parent's.
     */
    PermissionsManager.prototype.resolveNearestSecurableBaseUrl = function (artefact) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var current, base, _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        current = artefact;
                        _b.label = 1;
                    case 1:
                        base = this.buildBaseUrl(current);
                        _a = current.type === Permissions_1.SharePointArtefactType.Web;
                        if (_a) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.getHasUniqueRoleAssignments(base)];
                    case 2:
                        _a = (_b.sent());
                        _b.label = 3;
                    case 3:
                        if (_a) {
                            return [2 /*return*/, base];
                        }
                        current = current.type === Permissions_1.SharePointArtefactType.ListItem
                            ? { type: Permissions_1.SharePointArtefactType.List, webUrl: current.webUrl, listId: current.listId }
                            : { type: Permissions_1.SharePointArtefactType.Web, webUrl: current.webUrl };
                        _b.label = 4;
                    case 4: return [3 /*break*/, 1];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    PermissionsManager.prototype.getHasUniqueRoleAssignments = function (base) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, data;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.spHttpClient.get("".concat(base, "?$select=HasUniqueRoleAssignments"), sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 1:
                        response = _b.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _b.sent();
                        return [2 /*return*/, !!((_a = this.unwrapEntity(data)) === null || _a === void 0 ? void 0 : _a.HasUniqueRoleAssignments)];
                }
            });
        });
    };
    PermissionsManager.prototype.resolveLoginName = function (webUrl, loginHint) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var options, response, data, error_5;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        options = {
                            headers: {
                                'Accept': 'application/json;odata=verbose',
                                'Content-Type': 'application/json;odata=verbose'
                            },
                            body: JSON.stringify({ logonName: loginHint })
                        };
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, this.spHttpClient.post("".concat(this.normalizeWebUrl(webUrl), "/_api/web/ensureuser"), sp_http_1.SPHttpClient.configurations.v1, options)];
                    case 2:
                        response = _b.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _b.sent();
                        return [2 /*return*/, (_a = this.unwrapEntity(data)) === null || _a === void 0 ? void 0 : _a.LoginName];
                    case 4:
                        error_5 = _b.sent();
                        throw new Error("User could not be resolved on this web (".concat(loginHint, "): ").concat(error_5));
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /** Normalizes an odata=verbose entity response (`{ d: {...} }`) or a minimal/nometadata one (the object itself). */
    PermissionsManager.prototype.unwrapEntity = function (data) {
        var _a;
        return (_a = data === null || data === void 0 ? void 0 : data.d) !== null && _a !== void 0 ? _a : data;
    };
    /** Normalizes an odata=verbose collection response (`{ d: { results: [...] } }`) or a minimal/nometadata one (`{ value: [...] }`). */
    PermissionsManager.prototype.unwrapCollection = function (data) {
        var _a;
        if ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.results) {
            return data.d.results;
        }
        if (Array.isArray(data === null || data === void 0 ? void 0 : data.value)) {
            return data.value;
        }
        return [];
    };
    PermissionsManager.prototype.normalizeWebUrl = function (webUrl) {
        return webUrl.replace(/\/+$/, '');
    };
    PermissionsManager.prototype.toServerRelativeUrl = function (url) {
        if (!/^https?:\/\//i.test(url)) {
            return url;
        }
        var withoutProtocol = url.replace(/^https?:\/\//i, '');
        var firstSlash = withoutProtocol.indexOf('/');
        return firstSlash === -1 ? '/' : decodeURIComponent(withoutProtocol.substring(firstSlash));
    };
    PermissionsManager.prototype.buildBaseUrl = function (artefact) {
        var webUrl = this.normalizeWebUrl(artefact.webUrl);
        switch (artefact.type) {
            case Permissions_1.SharePointArtefactType.Web:
                return "".concat(webUrl, "/_api/web");
            case Permissions_1.SharePointArtefactType.List:
                if (!artefact.listId) {
                    throw new Error('artefact.listId is required for SharePointArtefactType.List');
                }
                return "".concat(webUrl, "/_api/web/lists('").concat(artefact.listId, "')");
            case Permissions_1.SharePointArtefactType.ListItem:
                if (!artefact.listId || artefact.itemId === undefined) {
                    throw new Error('artefact.listId and artefact.itemId are required for SharePointArtefactType.ListItem');
                }
                return "".concat(webUrl, "/_api/web/lists('").concat(artefact.listId, "')/items(").concat(artefact.itemId, ")");
            default:
                throw new Error("Unsupported SharePointArtefactType: ".concat(artefact.type));
        }
    };
    PermissionsManager.prototype.hasPermissionKind = function (mask, kind) {
        if (kind === Permissions_1.PermissionKind.FullMask) {
            return (mask.high & 0x7FFFFFFF) === 0x7FFFFFFF && mask.low === 0xFFFFFFFF;
        }
        var bit = kind - 1;
        if (bit < 32) {
            return (mask.low & (1 << bit)) !== 0;
        }
        return (mask.high & (1 << (bit - 32))) !== 0;
    };
    return PermissionsManager;
}());
exports.PermissionsManager = PermissionsManager;
exports.default = PermissionsManager;
//# sourceMappingURL=PermissionsManager.js.map