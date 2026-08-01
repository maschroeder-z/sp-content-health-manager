"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PermissionsManager = exports.ExpectedPermissionResolutionError = void 0;
var tslib_1 = require("tslib");
var sp_http_1 = require("@microsoft/sp-http");
var Permissions_1 = require("../models/REST/Permissions");
/**
 * Marks an error as an expected, already-explained condition (e.g. an Entra claim that turned
 * out to be a directory role rather than a group) so callers can log it quietly instead of as a
 * console.error - the message is already user-facing and actionable, not a bug to investigate.
 */
var ExpectedPermissionResolutionError = /** @class */ (function (_super) {
    tslib_1.__extends(ExpectedPermissionResolutionError, _super);
    function ExpectedPermissionResolutionError() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return ExpectedPermissionResolutionError;
}(Error));
exports.ExpectedPermissionResolutionError = ExpectedPermissionResolutionError;
var PermissionsManager = /** @class */ (function () {
    function PermissionsManager(msGraphClientFactory, spHttpClient) {
        this.spHttpClient = spHttpClient;
        this.graphClientPromise = msGraphClientFactory.getClient('3');
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
                                if (!member.LoginName && !member.Title) {
                                    console.warn('get4ArtefactPermissions: a role assignment\'s Member came back empty - the $expand=Member may have partially failed for this principal.', roleAssignment);
                                }
                                return {
                                    principalId: member.Id,
                                    principalType: principalType,
                                    isGroup: principalType !== Permissions_1.PrincipalType.User,
                                    displayName: member.Title || member.LoginName || '(unresolved principal)',
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
            var loginHint, loginName, error_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        loginHint = user.mail || user.userPrincipalName;
                        if (!loginHint) {
                            throw new Error('User must have a mail or userPrincipalName to resolve a SharePoint login name.');
                        }
                        return [4 /*yield*/, this.resolveLoginName(artefact.webUrl, loginHint)];
                    case 1:
                        loginName = _a.sent();
                        return [4 /*yield*/, this.getEffectivePermissions(loginName, artefact)];
                    case 2: return [2 /*return*/, _a.sent()];
                    case 3:
                        error_2 = _a.sent();
                        console.error('Error checking user permission:', error_2);
                        throw error_2;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Resolves whether an arbitrary principal (found via a search UI) has effective access to an
     * artefact, and if so, its effective capability breakdown.
     */
    PermissionsManager.prototype.checkAccess4Principal = function (principal, artefact) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var loginName, _a, permissionInfo, hasAccess, error_3;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 5, , 6]);
                        if (!/^\d+$/.test(principal.id)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.resolveSharePointGroupLoginName(artefact.webUrl, Number(principal.id))];
                    case 1:
                        _a = _b.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        _a = principal.id;
                        _b.label = 3;
                    case 3:
                        loginName = _a;
                        return [4 /*yield*/, this.getEffectivePermissions(loginName, artefact)];
                    case 4:
                        permissionInfo = _b.sent();
                        hasAccess = permissionInfo.rawMask.high !== 0 || permissionInfo.rawMask.low !== 0;
                        return [2 /*return*/, { hasAccess: hasAccess, permissionInfo: permissionInfo }];
                    case 5:
                        error_3 = _b.sent();
                        console.error('Error checking principal access:', error_3);
                        throw error_3;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    PermissionsManager.prototype.getEffectivePermissions = function (loginName, artefact) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var base, encodedLoginName, response, data, entity, mask;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        base = this.buildBaseUrl(artefact);
                        encodedLoginName = encodeURIComponent("'".concat(loginName.replace(/'/g, "''"), "'"));
                        return [4 /*yield*/, this.spHttpClient.get("".concat(base, "/getusereffectivepermissions(@v)?@v=").concat(encodedLoginName), sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
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
                }
            });
        });
    };
    PermissionsManager.prototype.resolveSharePointGroupLoginName = function (webUrl, groupId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, data, loginName;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.spHttpClient.get("".concat(this.normalizeWebUrl(webUrl), "/_api/web/sitegroups/getbyid(").concat(groupId, ")?$select=LoginName"), sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 1:
                        response = _b.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _b.sent();
                        loginName = (_a = this.unwrapEntity(data)) === null || _a === void 0 ? void 0 : _a.LoginName;
                        if (!loginName) {
                            throw new Error("Could not resolve LoginName for SharePoint group ".concat(groupId, "."));
                        }
                        return [2 /*return*/, loginName];
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
            var normalizedWebUrl, serverRelativeUrl, encodedServerRelativeUrl, response, data, entity, listId, itemId, error_4;
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
                        error_4 = _b.sent();
                        console.error('Error resolving artefact from file URL:', error_4);
                        throw error_4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Retrieves the extra per-page status shown by the Pages overview list's "Load details" action:
     * whether the page has an unpublished draft, whether it has unique (non-inherited) permissions,
     * and who (if anyone) has it checked out. Fetched in a single REST call by extending the same
     * GetFileByServerRelativeUrl/ListItemAllFields query shape used by resolveArtefactFromFileUrl.
     */
    PermissionsManager.prototype.getPageStatus = function (webUrl, fileUrl) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var normalizedWebUrl, serverRelativeUrl, encodedServerRelativeUrl, response, data, entity, file, checkedOutBy, error_5;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        normalizedWebUrl = this.normalizeWebUrl(webUrl);
                        serverRelativeUrl = this.toServerRelativeUrl(fileUrl);
                        encodedServerRelativeUrl = serverRelativeUrl
                            .replace(/'/g, "''")
                            .replace(/\(/g, '%28')
                            .replace(/\)/g, '%29');
                        return [4 /*yield*/, this.spHttpClient.get("".concat(normalizedWebUrl, "/_api/web/GetFileByServerRelativeUrl('").concat(encodedServerRelativeUrl, "')/ListItemAllFields") +
                                "?$select=HasUniqueRoleAssignments,File/Level,File/CheckOutType,File/CheckedOutByUser/Title" +
                                "&$expand=File,File/CheckedOutByUser", sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        entity = this.unwrapEntity(data);
                        file = entity === null || entity === void 0 ? void 0 : entity.File;
                        checkedOutBy = (file && file.CheckOutType !== 2 && file.CheckedOutByUser)
                            ? file.CheckedOutByUser.Title
                            : null;
                        return [2 /*return*/, {
                                needsApproval: (file === null || file === void 0 ? void 0 : file.Level) === 1, //'Draft',
                                hasUniquePermission: !!(entity === null || entity === void 0 ? void 0 : entity.HasUniqueRoleAssignments),
                                checkedOutBy: checkedOutBy
                            }];
                    case 3:
                        error_5 = _a.sent();
                        console.error('Error retrieving page status:', error_5);
                        throw error_5;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Resolves the site's pages library (the "Site Pages" list, BaseTemplate 119 - WebPageLibrary) as a
     * List-type artefact, for checking permissions on the library itself rather than a single page.
     * Filtering by BaseTemplate rather than title keeps this locale-independent.
     */
    PermissionsManager.prototype.resolvePagesLibraryArtefact = function (webUrl) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var normalizedWebUrl, response, data, lists, listId, error_6;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        normalizedWebUrl = this.normalizeWebUrl(webUrl);
                        return [4 /*yield*/, this.spHttpClient.get("".concat(normalizedWebUrl, "/_api/web/lists?$filter=BaseTemplate eq 119&$select=Id&$top=1"), sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 1:
                        response = _b.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _b.sent();
                        lists = this.unwrapCollection(data);
                        listId = (_a = lists === null || lists === void 0 ? void 0 : lists[0]) === null || _a === void 0 ? void 0 : _a.Id;
                        if (!listId) {
                            throw new Error('Could not resolve the pages library for this site.');
                        }
                        return [2 /*return*/, {
                                type: Permissions_1.SharePointArtefactType.List,
                                webUrl: normalizedWebUrl,
                                listId: listId
                            }];
                    case 3:
                        error_6 = _b.sent();
                        console.error('Error resolving pages library artefact:', error_6);
                        throw error_6;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    PermissionsManager.prototype.hasUniquePermission = function (artefact) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var error_7;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getHasUniqueRoleAssignments(this.buildBaseUrl(artefact))];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2:
                        error_7 = _a.sent();
                        console.error('Error checking unique permissions:', error_7);
                        throw error_7;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Resolves all users belonging to a group principal, whether it is a native SharePoint group
     * (resolved via SharePoint REST) or an Entra ID group (resolved via Microsoft Graph, since Entra
     * groups are not queryable through SharePoint REST).
     */
    PermissionsManager.prototype.resolveUser4Group = function (groupInfo) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var entraGroupId, error_8;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        if (!(groupInfo.principalType === Permissions_1.PrincipalType.SharePointGroup)) return [3 /*break*/, 2];
                        if (!groupInfo.principalId) {
                            throw new Error('groupInfo.principalId is required to resolve a SharePoint group.');
                        }
                        return [4 /*yield*/, this.resolveSharePointGroupUsers(groupInfo.webUrl, groupInfo.principalId)];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2:
                        entraGroupId = this.tryExtractEntraGroupId(groupInfo.loginName);
                        if (!entraGroupId) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.resolveEntraGroupUsers(entraGroupId)];
                    case 3: return [2 /*return*/, _a.sent()];
                    case 4: throw new Error("Unsupported group principal (not a SharePoint group or Entra-backed security group): ".concat(groupInfo.loginName || groupInfo.displayName));
                    case 5:
                        error_8 = _a.sent();
                        if (error_8 instanceof ExpectedPermissionResolutionError) {
                            console.warn('Could not resolve group users:', error_8.message);
                        }
                        else {
                            console.error('Error resolving group users:', error_8);
                        }
                        throw error_8;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Resolves the direct child groups of a group principal (one level, not transitive) — unlike
     * resolveUser4Group, which resolves effective/transitive users and does not surface nested-group
     * structure. Used to lazily populate a group hierarchy (e.g. a tree UI) one expansion at a time.
     */
    PermissionsManager.prototype.resolveNestedGroups = function (groupInfo) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var entraGroupId, error_9;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        if (!(groupInfo.principalType === Permissions_1.PrincipalType.SharePointGroup)) return [3 /*break*/, 2];
                        if (!groupInfo.principalId) {
                            throw new Error('groupInfo.principalId is required to resolve a SharePoint group.');
                        }
                        return [4 /*yield*/, this.resolveSharePointGroupNestedGroups(groupInfo.webUrl, groupInfo.principalId)];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2:
                        entraGroupId = this.tryExtractEntraGroupId(groupInfo.loginName);
                        if (!entraGroupId) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.resolveEntraGroupNestedGroups(groupInfo.webUrl, entraGroupId)];
                    case 3: return [2 /*return*/, _a.sent()];
                    case 4: throw new Error("Unsupported group principal (not a SharePoint group or Entra-backed security group): ".concat(groupInfo.loginName || groupInfo.displayName));
                    case 5:
                        error_9 = _a.sent();
                        if (error_9 instanceof ExpectedPermissionResolutionError) {
                            console.warn('Could not resolve nested groups:', error_9.message);
                        }
                        else {
                            console.error('Error resolving nested groups:', error_9);
                        }
                        throw error_9;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Resolves the members of a built-in Entra directory role (e.g. Global Administrator), given its
     * universal role template id (see models/REST/Permissions.ts SHAREPOINT_RELEVANT_ENTRA_ROLES).
     * Unlike security groups, directory roles are addressed via /directoryRoles, not /groups - a role
     * that has never been assigned to anyone in this tenant may not be "activated" yet and this call
     * 404s for it (activating one requires the write-scoped RoleManagement.ReadWrite.Directory
     * permission, out of scope for a read-only permissions audit tool). That 404 is treated as "nobody
     * holds this role" rather than an error and resolves to an empty list. Requires the
     * RoleManagement.Read.Directory Graph permission, requested in config/package-solution.json and
     * subject to tenant admin approval - until approved this throws a 403 GraphError.
     */
    PermissionsManager.prototype.resolveDirectoryRoleUsers = function (roleTemplateId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var client, role, error_10;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        return [4 /*yield*/, this.graphClientPromise];
                    case 1:
                        client = _a.sent();
                        return [4 /*yield*/, client
                                .api("/directoryRoles(roleTemplateId='".concat(encodeURIComponent(roleTemplateId), "')"))
                                .version('v1.0')
                                .select('id')
                                .get()];
                    case 2:
                        role = _a.sent();
                        return [4 /*yield*/, this.fetchDirectoryRoleMembersByRoleId(client, role.id)];
                    case 3: return [2 /*return*/, _a.sent()];
                    case 4:
                        error_10 = _a.sent();
                        if ((error_10 === null || error_10 === void 0 ? void 0 : error_10.statusCode) === 404) {
                            return [2 /*return*/, []];
                        }
                        console.error('Error resolving directory role members:', error_10);
                        throw error_10;
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Fetches the members of an already-activated directory role, given its own (tenant-specific)
     * object id - not a roleTemplateId. Shared by resolveDirectoryRoleUsers (which resolves a
     * universal roleTemplateId to this id first) and resolveEntraGroupUsers's directory-role
     * fallback (which already has what's presumably this id, extracted directly from a SharePoint
     * claim, so it skips the roleTemplateId resolution step entirely).
     */
    PermissionsManager.prototype.fetchDirectoryRoleMembersByRoleId = function (client, roleId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, members;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, client
                            .api("/directoryRoles/".concat(encodeURIComponent(roleId), "/members"))
                            .version('v1.0')
                            .select(['id', 'displayName', 'mail', 'userPrincipalName'].join(','))
                            .get()];
                    case 1:
                        response = _a.sent();
                        members = (response === null || response === void 0 ? void 0 : response.value) || [];
                        return [2 /*return*/, members.map(function (member) { return ({
                                id: member.id,
                                displayName: member.displayName,
                                email: member.mail || member.userPrincipalName || undefined
                            }); })];
                }
            });
        });
    };
    PermissionsManager.prototype.resolveSharePointGroupUsers = function (webUrl, groupId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var members;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.fetchSharePointGroupMembers(webUrl, groupId)];
                    case 1:
                        members = _a.sent();
                        return [2 /*return*/, members
                                .filter(function (member) { return !_this.tryExtractEntraGroupId(member.LoginName); })
                                .map(function (member) { return ({
                                id: String(member.Id),
                                displayName: member.Title || member.LoginName || '(unresolved principal)',
                                email: member.Email || undefined,
                                loginName: member.LoginName
                            }); })];
                }
            });
        });
    };
    /**
     * A SharePoint group cannot contain another SharePoint group, but an Entra group can be added as a
     * member of one — it then shows up in the group's `users` collection as a `c:0t.c|tenant|<guid>`
     * login name. SharePoint's legacy `sitegroups(id)/users` endpoint reports this shadow entity's
     * PrincipalType as a plain User, so the login-name claims pattern (not PrincipalType) is the only
     * reliable way to recognize it here.
     */
    PermissionsManager.prototype.resolveSharePointGroupNestedGroups = function (webUrl, groupId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var members;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.fetchSharePointGroupMembers(webUrl, groupId)];
                    case 1:
                        members = _a.sent();
                        return [2 /*return*/, members
                                .filter(function (member) { return !!_this.tryExtractEntraGroupId(member.LoginName); })
                                .map(function (member) { return ({
                                webUrl: webUrl,
                                principalId: member.Id,
                                principalType: Permissions_1.PrincipalType.SecurityGroup,
                                loginName: member.LoginName,
                                displayName: member.Title
                            }); })];
                }
            });
        });
    };
    /**
     * SharePoint REST collections default to a ~100-item page unless $top is specified, and this
     * endpoint doesn't automatically follow up - without pagination, groups larger than that would
     * silently lose members past the first page. $top=5000 matches SharePoint's own hard collection
     * cap, and d.__next (odata=verbose) is followed in case a tenant enforces a smaller page size.
     */
    PermissionsManager.prototype.fetchSharePointGroupMembers = function (webUrl, groupId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var members, nextUrl, response, data, missingDetails;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        members = [];
                        nextUrl = "".concat(this.normalizeWebUrl(webUrl), "/_api/web/sitegroups/getbyid(").concat(groupId, ")/users?$top=5000");
                        _b.label = 1;
                    case 1:
                        if (!nextUrl) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.spHttpClient.get(nextUrl, sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 2:
                        response = _b.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _b.sent();
                        members.push.apply(members, this.unwrapCollection(data));
                        nextUrl = ((_a = data === null || data === void 0 ? void 0 : data.d) === null || _a === void 0 ? void 0 : _a.__next) || undefined;
                        return [3 /*break*/, 1];
                    case 4:
                        missingDetails = members.filter(function (member) { return !member.LoginName || !member.Title; });
                        if (missingDetails.length > 0) {
                            console.warn("fetchSharePointGroupMembers: ".concat(missingDetails.length, " member(s) of group ").concat(groupId, " came back with a missing LoginName/Title - the REST expand may have partially failed for these principals."), missingDetails);
                        }
                        return [2 /*return*/, members];
                }
            });
        });
    };
    PermissionsManager.prototype.resolveEntraGroupUsers = function (groupId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var client, response, members, error_11, roleError_1;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.graphClientPromise];
                    case 1:
                        client = _a.sent();
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 4, , 9]);
                        return [4 /*yield*/, client
                                .api("/groups/".concat(encodeURIComponent(groupId), "/transitiveMembers/microsoft.graph.user"))
                                .version('v1.0')
                                .select(['id', 'displayName', 'mail', 'userPrincipalName'].join(','))
                                .get()];
                    case 3:
                        response = _a.sent();
                        members = (response === null || response === void 0 ? void 0 : response.value) || [];
                        return [2 /*return*/, members.map(function (member) { return ({
                                id: member.id,
                                displayName: member.displayName,
                                email: member.mail || member.userPrincipalName || undefined
                            }); })];
                    case 4:
                        error_11 = _a.sent();
                        if (!((error_11 === null || error_11 === void 0 ? void 0 : error_11.statusCode) === 404)) return [3 /*break*/, 8];
                        _a.label = 5;
                    case 5:
                        _a.trys.push([5, 7, , 8]);
                        return [4 /*yield*/, this.fetchDirectoryRoleMembersByRoleId(client, groupId)];
                    case 6: return [2 /*return*/, _a.sent()];
                    case 7:
                        roleError_1 = _a.sent();
                        if ((roleError_1 === null || roleError_1 === void 0 ? void 0 : roleError_1.statusCode) === 403) {
                            throw new ExpectedPermissionResolutionError("This Entra ID object (".concat(groupId, ") appears to be a built-in directory role, but your tenant admin has not approved this app's RoleManagement.Read.Directory permission request yet - approve it in the SharePoint Admin Center's API access page to list its members."));
                        }
                        return [3 /*break*/, 8];
                    case 8: throw this.toExpectedGraphGroupError(error_11, groupId);
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * SharePoint represents both real Entra security groups and built-in Entra directory roles
     * (Global Administrator, SharePoint Administrator, etc.) with the same claim format
     * (`c:0t.c|tenant|<guid>`), so a claim that's actually a role id (or a since-deleted group)
     * 404s here - Graph's /groups endpoint doesn't recognize either as a group. Surface that as a
     * clear, expected condition instead of letting Graph's raw error propagate as a console-error
     * storm; a 403 means the same claim resolved but this app isn't authorized to read it (e.g. a
     * role membership read that needs a permission this app hasn't been granted yet).
     */
    PermissionsManager.prototype.toExpectedGraphGroupError = function (error, groupId) {
        var statusCode = error === null || error === void 0 ? void 0 : error.statusCode;
        if (statusCode === 404) {
            return new ExpectedPermissionResolutionError("This Entra ID object (".concat(groupId, ") couldn't be found as a security group - it may be a built-in directory role (like Global Administrator), which SharePoint represents using the same claim format, or a security group that has since been deleted. Role membership can't be listed here without the RoleManagement.Read.Directory permission."));
        }
        if (statusCode === 403) {
            return new ExpectedPermissionResolutionError("Access denied reading Entra ID object ".concat(groupId, " - your tenant admin has not approved this app's pending Graph API permission request yet. Approve it in the SharePoint Admin Center's API access page to continue."));
        }
        return error instanceof Error ? error : new Error(String(error));
    };
    /** Direct (non-transitive) nested groups of an Entra group, one level down. */
    PermissionsManager.prototype.resolveEntraGroupNestedGroups = function (webUrl, groupId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var client, error_12, roleError_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.graphClientPromise];
                    case 1:
                        client = _a.sent();
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 4, , 9]);
                        return [4 /*yield*/, this.fetchEntraGroupNestedGroups(client, webUrl, groupId)];
                    case 3: return [2 /*return*/, _a.sent()];
                    case 4:
                        error_12 = _a.sent();
                        if (!((error_12 === null || error_12 === void 0 ? void 0 : error_12.statusCode) === 404)) return [3 /*break*/, 8];
                        _a.label = 5;
                    case 5:
                        _a.trys.push([5, 7, , 8]);
                        return [4 /*yield*/, this.fetchDirectoryRoleNestedGroups(client, webUrl, groupId)];
                    case 6: return [2 /*return*/, _a.sent()];
                    case 7:
                        roleError_2 = _a.sent();
                        if ((roleError_2 === null || roleError_2 === void 0 ? void 0 : roleError_2.statusCode) === 403) {
                            throw new ExpectedPermissionResolutionError("This Entra ID object (".concat(groupId, ") appears to be a built-in directory role, but your tenant admin has not approved this app's RoleManagement.Read.Directory permission request yet - approve it in the SharePoint Admin Center's API access page to list its nested groups."));
                        }
                        return [3 /*break*/, 8];
                    case 8: throw this.toExpectedGraphGroupError(error_12, groupId);
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    PermissionsManager.prototype.fetchEntraGroupNestedGroups = function (client, webUrl, groupId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, members;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, client
                            .api("/groups/".concat(encodeURIComponent(groupId), "/members/microsoft.graph.group"))
                            .version('v1.0')
                            .select(['id', 'displayName'].join(','))
                            .get()];
                    case 1:
                        response = _a.sent();
                        members = (response === null || response === void 0 ? void 0 : response.value) || [];
                        return [2 /*return*/, members.map(function (member) { return ({
                                webUrl: webUrl,
                                principalType: Permissions_1.PrincipalType.SecurityGroup,
                                loginName: "c:0t.c|tenant|".concat(member.id),
                                displayName: member.displayName
                            }); })];
                }
            });
        });
    };
    /**
     * Directory roles are usually assigned to users, but Entra RBAC also allows assigning a role to a
     * group (its members then inherit the role) - so, mirroring fetchEntraGroupNestedGroups, this
     * asks Graph for the group-typed members of an already-activated role, given its own object id.
     */
    PermissionsManager.prototype.fetchDirectoryRoleNestedGroups = function (client, webUrl, roleId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var response, members;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, client
                            .api("/directoryRoles/".concat(encodeURIComponent(roleId), "/members/microsoft.graph.group"))
                            .version('v1.0')
                            .select(['id', 'displayName'].join(','))
                            .get()];
                    case 1:
                        response = _a.sent();
                        members = (response === null || response === void 0 ? void 0 : response.value) || [];
                        return [2 /*return*/, members.map(function (member) { return ({
                                webUrl: webUrl,
                                principalType: Permissions_1.PrincipalType.SecurityGroup,
                                loginName: "c:0t.c|tenant|".concat(member.id),
                                displayName: member.displayName
                            }); })];
                }
            });
        });
    };
    /** Extracts the AAD group id from a claims-encoded Entra group login name, e.g. `c:0t.c|tenant|<guid>`. */
    PermissionsManager.prototype.tryExtractEntraGroupId = function (loginName) {
        if (!loginName) {
            return undefined;
        }
        var match = /^c:0t\.c\|tenant\|([0-9a-fA-F-]{36})$/.exec(loginName);
        return match ? match[1] : undefined;
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
            var options, response, data, error_13;
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
                        error_13 = _b.sent();
                        throw new Error("User could not be resolved on this web (".concat(loginHint, "): ").concat(error_13));
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