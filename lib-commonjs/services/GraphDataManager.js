"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.GraphDataManager = void 0;
var tslib_1 = require("tslib");
var sp_http_1 = require("@microsoft/sp-http");
var GraphDataManager = /** @class */ (function () {
    function GraphDataManager(msGraphClientFactory, spHttpClient) {
        this.graphClientPromise = msGraphClientFactory.getClient('3');
        this.spHTTPClient = spHttpClient;
    }
    // ?$select=webUrl,Guid&$filter=siteCollection/root%20ne%20null
    /*public async GetSites(parentSite?: Site): Promise<Site[]> {
      const client = await this.graphClientPromise;
  
      if (parentSite?.id) {
        const response = await client
          .api(`/sites/${encodeURIComponent(parentSite.id)}/sites`)
          .version('v1.0')
          .select(['id', 'name', 'displayName', 'webUrl', 'siteCollection'].join(','))
          .get();
  
        const items: Site[] = (response?.value || []).map((s: any) => ({
          id: s.id,
          name: s.name,
          displayName: s.displayName,
          webUrl: s.webUrl,
          siteCollection: s.siteCollection
        }));
        return items;
      }
  
      // Top-level site collections: search all sites, then keep those with siteCollection present
      const searchResponse = await client
        .api('/sites/getAllSites')
        .version('v1.0')
        .select(['id', 'name', 'displayName', 'webUrl', 'siteCollection'].join(','))
        .get();
  
      const allSites: Site[] = (searchResponse?.value || []).map((s: any) => ({
        id: s.id,
        name: s.name,
        displayName: s.displayName,
        webUrl: s.webUrl,
        siteCollection: s.siteCollection
      }));
  
      const topLevelSites = allSites.filter(s => !!s.siteCollection);
      return topLevelSites;
    }*/
    // https://learn.microsoft.com/en-us/graph/api/resources/sitepage?view=graph-rest-1.0
    GraphDataManager.prototype.GetPageContent = function (siteID, pageID) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var client, response;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.graphClientPromise];
                    case 1:
                        client = _a.sent();
                        return [4 /*yield*/, client
                                .api("/sites/".concat(encodeURIComponent(siteID), "/pages/").concat(pageID, "/microsoft.graph.sitePage?$expand=canvasLayout"))
                                .version('v1.0')
                                .select(['id', 'name', 'title', 'webUrl', 'createdDateTime', 'lastModifiedDateTime'].join(','))
                                .get()];
                    case 2:
                        response = _a.sent();
                        return [2 /*return*/, response];
                }
            });
        });
    };
    GraphDataManager.prototype.GetPages4Site = function (siteID) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var client, response, items;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.graphClientPromise];
                    case 1:
                        client = _a.sent();
                        return [4 /*yield*/, client
                                .api("/sites/".concat(encodeURIComponent(siteID), "/pages/microsoft.graph.sitePage"))
                                .version('v1.0')
                                .select(['id', 'name', 'title', 'webUrl', 'createdDateTime', 'lastModifiedDateTime'].join(','))
                                .get()];
                    case 2:
                        response = _a.sent();
                        items = ((response === null || response === void 0 ? void 0 : response.value) || []).map(function (p) { return ({
                            id: p.id,
                            name: p.name,
                            title: p.title,
                            webUrl: p.webUrl,
                            createdDateTime: p.createdDateTime,
                            lastModifiedDateTime: p.lastModifiedDateTime,
                            InProgress: false
                        }); });
                        return [2 /*return*/, items];
                }
            });
        });
    };
    GraphDataManager.prototype.GetLibraries = function (siteID) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var client, response;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.graphClientPromise];
                    case 1:
                        client = _a.sent();
                        return [4 /*yield*/, client
                                .api("/sites/".concat(encodeURIComponent(siteID), "/lists"))
                                .version('v1.0')
                                .select(['id', 'name', 'displayName', 'webUrl', 'createdDateTime', 'lastModifiedDateTime'].join(','))
                                .get()];
                    case 2:
                        response = _a.sent();
                        return [2 /*return*/, response.value];
                }
            });
        });
    };
    GraphDataManager.prototype.GetAllLists = function (siteUrl, incLists, incLibraries) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var apiUrl, response, data, lists, error_1;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        apiUrl = "".concat(siteUrl, "/_api/web/lists?$expand=DefaultView");
                        return [4 /*yield*/, fetch(apiUrl, {
                                method: 'GET',
                                headers: {
                                    'Accept': 'application/json;odata=verbose',
                                    'Content-Type': 'application/json'
                                },
                                credentials: 'include' // Include cookies for authentication
                            })];
                    case 1:
                        response = _b.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _b.sent();
                        lists = ((_a = data.d) === null || _a === void 0 ? void 0 : _a.results.filter(function (x) { return (x.BaseType === 0 && incLists) || (x.BaseTemplate === 101 && x.BaseType === 1 && incLibraries); })) || [];
                        return [2 /*return*/, lists.map(function (list) { return ({
                                AllowContentTypes: list.AllowContentTypes,
                                BaseTemplate: list.BaseTemplate,
                                BaseType: list.BaseType,
                                ContentTypesEnabled: list.ContentTypesEnabled,
                                CrawlNonDefaultViews: list.CrawlNonDefaultViews,
                                Created: list.Created,
                                CurrentChangeToken: list.CurrentChangeToken,
                                DefaultContentApprovalWorkflowId: list.DefaultContentApprovalWorkflowId,
                                DefaultItemOpenUseListSetting: list.DefaultItemOpenUseListSetting,
                                Description: list.Description,
                                Direction: list.Direction,
                                DisableCommenting: list.DisableCommenting,
                                DisableGridEditing: list.DisableGridEditing,
                                DocumentTemplateUrl: list.DocumentTemplateUrl,
                                DraftVersionVisibility: list.DraftVersionVisibility,
                                EnableAttachments: list.EnableAttachments,
                                EnableFolderCreation: list.EnableFolderCreation,
                                EnableMinorVersions: list.EnableMinorVersions,
                                EnableModeration: list.EnableModeration,
                                EnableRequestSignOff: list.EnableRequestSignOff,
                                EnableVersioning: list.EnableVersioning,
                                EntityTypeName: list.EntityTypeName,
                                ExemptFromBlockDownloadOfNonViewableFiles: list.ExemptFromBlockDownloadOfNonViewableFiles,
                                FileSavePostProcessingEnabled: list.FileSavePostProcessingEnabled,
                                ForceCheckout: list.ForceCheckout,
                                HasExternalDataSource: list.HasExternalDataSource,
                                Hidden: list.Hidden,
                                Id: list.Id,
                                ImagePath: list.ImagePath,
                                ImageUrl: list.ImageUrl,
                                DefaultSensitivityLabelForLibrary: list.DefaultSensitivityLabelForLibrary,
                                SensitivityLabelToEncryptOnDownloadForLibrary: list.SensitivityLabelToEncryptOnDownloadForLibrary,
                                IrmEnabled: list.IrmEnabled,
                                IrmExpire: list.IrmExpire,
                                IrmReject: list.IrmReject,
                                IsApplicationList: list.IsApplicationList,
                                IsCatalog: list.IsCatalog,
                                IsPrivate: list.IsPrivate,
                                ItemCount: list.ItemCount,
                                LastItemDeletedDate: list.LastItemDeletedDate,
                                LastItemModifiedDate: list.LastItemModifiedDate,
                                LastItemUserModifiedDate: list.LastItemUserModifiedDate,
                                ListExperienceOptions: list.ListExperienceOptions,
                                ListItemEntityTypeFullName: list.ListItemEntityTypeFullName,
                                MajorVersionLimit: list.MajorVersionLimit,
                                MajorWithMinorVersionsLimit: list.MajorWithMinorVersionsLimit,
                                MultipleDataList: list.MultipleDataList,
                                NoCrawl: list.NoCrawl,
                                ParentWebPath: list.ParentWebPath,
                                ParserDisabled: list.ParserDisabled,
                                ServerTemplateCanCreateFolders: list.ServerTemplateCanCreateFolders,
                                TemplateFeatureId: list.TemplateFeatureId,
                                Title: list.Title,
                                DefaultView: list.DefaultView,
                                ParentWebUrl: list.ParentWebUrl + "/" + list.EntityTypeName
                            }); })];
                    case 3:
                        error_1 = _b.sent();
                        console.error('Error fetching lists:', error_1);
                        throw error_1;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Queries checked-out items using classic SharePoint REST instead of Graph.
     * Graph's /items endpoint rejects $filter on person-field-derived names like
     * CheckoutUserLookupId ("A provided field name is not recognized"). Classic REST also
     * rejects $select/$expand=CheckoutUser ("field or property does not exist") - the
     * checkout person field's navigation property is actually named CheckoutUserId (the "Id"
     * suffix is part of the nav property name here, not just the lookup id column), and
     * unlike File/CheckedOutByUser it filters and expands directly on the /items collection.
     */
    GraphDataManager.prototype.Query4CheckedOutItems = function (site, listID, defaultUrl, dateStart) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var apiUrl, response, data, items, error_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        defaultUrl = site.url + "/_layouts/15/listform.aspx?PageType=4&ListId=";
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        apiUrl = "".concat(site.url, "/_api/web/lists('").concat(listID, "')/items") +
                            "?$select=Id,FileLeafRef,Created,Modified,ContentTypeId,CheckoutUser/Title,CheckoutUser/EMail" +
                            "&$expand=CheckoutUser,ContentType" +
                            "&$filter=CheckoutUserId ne null";
                        return [4 /*yield*/, this.spHTTPClient.get(apiUrl, sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        console.log("check-out items:", data);
                        items = (data.value || []).map(function (item) {
                            var _a;
                            return (tslib_1.__assign(tslib_1.__assign({}, item), { Title: item.FileLeafRef, CheckedOutBy: ((_a = item.CheckoutUser) === null || _a === void 0 ? void 0 : _a.Title) || null, webUrl: "".concat(defaultUrl).concat(listID, "&id=").concat(item.Id) }));
                        });
                        return [2 /*return*/, items];
                    case 4:
                        error_2 = _a.sent();
                        console.error('Error querying checked-out items:', error_2);
                        throw error_2;
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    GraphDataManager.prototype.Query4ItemByDate = function (site, listID, defaultUrl, dateStart) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var client, formattedDate, urlToDetails, urlDispForm_1, response, items, error_3;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.graphClientPromise];
                    case 1:
                        client = _a.sent();
                        formattedDate = dateStart.toISOString();
                        urlToDetails = defaultUrl.split("/");
                        urlToDetails.pop();
                        urlDispForm_1 = urlToDetails.join("/") + "/_layouts/15/listform.aspx?PageType=4";
                        return [4 /*yield*/, client
                                .api("/sites/".concat(encodeURIComponent(site.id), "/lists/").concat(listID, "/items"))
                                .version('v1.0')
                                .filter("fields/Modified le '".concat(formattedDate, "'"))
                                .expand('fields')
                                .select(['id', 'fields'])
                                .get()];
                    case 2:
                        response = _a.sent();
                        items = ((response === null || response === void 0 ? void 0 : response.value) || []).map(function (item) { return (tslib_1.__assign(tslib_1.__assign({ Id: item.id, Title: item.fields.Title || item.fields.FileLeafRef, Created: item.fields.Created, Modified: item.fields.Modified, ContentTypeId: item.fields.ContentTypeId }, item.fields), { webUrl: "".concat(urlDispForm_1, "&id=").concat(item.id, "&listid=").concat(listID) })); });
                        return [2 /*return*/, items];
                    case 3:
                        error_3 = _a.sent();
                        console.error('Error querying items by date:', error_3);
                        throw error_3;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /* NOTE: The following method is commented out because it uses the Microsoft Graph API to get permissions for a list item,
             but the Graph API may not support this operation for all scenarios. If you need to retrieve permissions, consider
             using SharePoint REST API or other methods.
      public async GetPermission4Item(site: Site, listID: string, listItemID: string): Promise<MicrosoftGraphBeta.Permission[]> {
      try {
        const client = await this.graphClientPromise;
        // Query for permission information using Microsoft Graph API
        const response = await client
          .api(`/sites/${encodeURIComponent(site.id)}/lists/${listID}/items/${listItemID}/permissions`)
          .version('beta')
          .get();
        console.log(response?.value);
        return response?.value || [];
      } catch (error) {
        console.error('Error retrieving item permissions:', error);
        throw error;
      }
    }*/
    /**
   * Queries list items by date using SharePoint REST API
   * Endpoint: /[siteUrl]/_api/web/lists('[listID]')/GetItems(query=@v1)?@v1={'ViewXml':'<View><Query><Where><Leq><FieldRef Name=Modified/><Value Type=DateTime>[dateStart]</Value></Leq></Where></Query></View>'}&$expand=file
   */
    GraphDataManager.prototype.Query4ItemByDateClassic = function (siteUrl, listID, defaultUrl, dateStart) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var formattedDate, viewXml, options, apiUrl, response, data, items, error_4;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!(typeof defaultUrl !== "undefined")) return [3 /*break*/, 5];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 4, , 5]);
                        formattedDate = dateStart.toISOString();
                        // /sites/Demo02/Freigegebene Dokumente/Forms/AllItems.aspx /sites/Demo02/FormServerTemplates/Forms/All Forms.aspx
                        /*const temp = defaultUrl.split("/")
                        temp.pop();
                        //temp.push("ViewForm.aspx?id=");
                        temp.push("_layouts/15/listform.aspx?PageType=4&ListId=");
                        defaultUrl = temp.join("/");*/
                        //https://plumsail.com/docs/forms-sp/how-to/link-to-form.html
                        defaultUrl = siteUrl + "/_layouts/15/listform.aspx?PageType=4&ListId=";
                        viewXml = "<View><Query><Where><Leq><FieldRef Name=Modified/><Value Type=DateTime>".concat(formattedDate, "</Value></Leq></Where></Query></View>");
                        options = {
                            headers: {
                                'odata-version': '3.0',
                                'Accept': 'application/json;odata=verbose',
                                'Content-Type': 'application/json'
                            },
                            body: "{'query': {          \n            'ViewXml':'".concat(viewXml, "'\n          }}")
                        };
                        apiUrl = "".concat(siteUrl, "/_api/web/lists('").concat(listID, "')/GetItems?$expand=ParentList,File,ContentType");
                        return [4 /*yield*/, this.spHTTPClient.post(apiUrl, sp_http_1.SPHttpClient.configurations.v1, options)];
                    case 2:
                        response = _b.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _b.sent();
                        items = ((_a = data.d) === null || _a === void 0 ? void 0 : _a.results) || [];
                        items.forEach(function (item) {
                            // https://[Your SharePoint SiteURL]/_layouts/15/listform.aspx?PageType=[Type]&ListId=[ListGUID]&ID=[Item ID]
                            //console.log(item);          
                            item.webUrl = "".concat(defaultUrl).concat(item.ParentList.Id, "&id=").concat(item.Id);
                            //item.webUrl = `/_layouts/15/listform.aspx?PageType=4&ListId=${(item as any).GUID}`
                        });
                        return [2 /*return*/, items];
                    case 4:
                        error_4 = _b.sent();
                        console.error('Error querying items by date:', error_4);
                        throw error_4;
                    case 5: return [2 /*return*/, []];
                }
            });
        });
    };
    return GraphDataManager;
}());
exports.GraphDataManager = GraphDataManager;
exports.default = GraphDataManager;
//# sourceMappingURL=GraphDataManager.js.map