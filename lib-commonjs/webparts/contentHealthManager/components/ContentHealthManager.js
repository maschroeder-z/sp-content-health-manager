"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var ContentHealthManager_module_scss_1 = tslib_1.__importDefault(require("./ContentHealthManager.module.scss"));
var ListView_1 = require("@pnp/spfx-controls-react/lib/ListView");
var react_1 = require("@fluentui/react");
var SitePicker_1 = require("@pnp/spfx-controls-react/lib/SitePicker");
var react_components_1 = require("@fluentui/react-components");
var react_resizable_panels_1 = require("react-resizable-panels");
var PeoplePicker_1 = require("@pnp/spfx-controls-react/lib/PeoplePicker");
var GraphDataManager_1 = tslib_1.__importDefault(require("../../../services/GraphDataManager"));
var PageProcessing_1 = require("../../../Core/PageProcessing");
var react_icons_1 = require("@fluentui/react-icons");
var PermissionsManager_1 = tslib_1.__importDefault(require("../../../services/PermissionsManager"));
var spfx_controls_react_1 = require("@pnp/spfx-controls-react");
var ListTemplateTypes_1 = require("../../../Core/ListTemplateTypes");
var strings = tslib_1.__importStar(require("ContentHealthManagerWebPartStrings"));
var ContentHealthManager = /** @class */ (function (_super) {
    tslib_1.__extends(ContentHealthManager, _super);
    function ContentHealthManager(props) {
        var _this = _super.call(this, props) || this;
        _this.tempSelectedSites = [
            {
                "id": "0a83c49d-6da8-459e-8bb4-98be06a28dcc",
                "webId": "ca9dc690-1f36-49b3-9283-05547458d435",
                "title": "Meine Schulung",
                "url": "https://devsky365.sharepoint.com/sites/Demo03"
            },
            {
                "id": "399408ed-462d-4ec4-acfd-69ee87b54649",
                "webId": "ca9dc690-1f36-49b3-9283-05547458d435",
                "title": "Make your own LOB :-)",
                "url": "https://devsky365.sharepoint.com/sites/my-own-lob-apps"
            },
            {
                "id": "15908e6d-d68a-4154-a9b7-a8557f5ace69",
                "webId": "ea4629cd-d579-48e8-9c74-9505c13fd042",
                "title": "HeimHaus",
                "url": "https://devsky365.sharepoint.com/sites/HeimHaus"
            },
            {
                "id": "d6f6d04c-5c5b-468c-82d7-39d08e86dfa5",
                "webId": "eb707bcc-5ead-49c5-81bc-3109c317f837",
                "title": "Hausfeen",
                "url": "https://devsky365.sharepoint.com/sites/Hausfeen"
            }
        ];
        // View fields for found items in library report dialog
        _this.viewFieldsFoundItems = [
            { name: 'Id', displayName: 'ID', sorting: true, isResizable: true, minWidth: 80, linkPropertyName: 'webUrl' },
            { name: 'Title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 200 },
            {
                name: 'Created', displayName: 'Created', sorting: true, isResizable: true, minWidth: 120,
                render: function (item, index, column) {
                    var date = new Date(item.Created);
                    return React.createElement(spfx_controls_react_1.FieldDateRenderer, { text: date.toLocaleDateString() });
                }
            },
            {
                name: 'Modified', displayName: 'Modified', sorting: true, isResizable: true, minWidth: 120,
                render: function (item, index, column) {
                    var date = new Date(item.Modified);
                    return React.createElement(spfx_controls_react_1.FieldDateRenderer, { text: date.toLocaleDateString() });
                }
            },
            { name: 'ContentTypeId', displayName: 'Content Type', sorting: true, isResizable: true, minWidth: 150,
                render: function (item, inxdex, column) {
                    if (typeof item.ContentType !== "undefined")
                        return item.ContentType;
                    return item["ContentType.Name"];
                }
            }
        ];
        // BaseTemplate BaseType EnableAttachments EnableFolderCreation EnableVersioning ForceCheckout ItemCount LastItemModifiedDate LastItemUserModifiedDate
        _this.viewFieldsLibs = [
            { name: 'Title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 120, linkPropertyName: 'DefaultView.ServerRelativeUrl' },
            { name: 'ItemCount', displayName: 'Items', sorting: true, isResizable: true, minWidth: 120 },
            { name: 'FoundItems', displayName: strings.FoundLabel, sorting: true, isResizable: true, minWidth: 120,
                render: function (item, index, column) {
                    var _a;
                    var entry = _this.GetLibraryEntryByIndex(item.Id);
                    if (typeof entry.FoundItems !== "undefined" && entry.FoundItems !== null) {
                        return React.createElement(spfx_controls_react_1.FieldTextRenderer, { text: "".concat(strings.FoundLabel, ": ").concat((_a = entry.FoundItems) === null || _a === void 0 ? void 0 : _a.length) });
                    }
                    else
                        return React.createElement(spfx_controls_react_1.FieldTextRenderer, { text: strings.StartQueryForResults });
                }
            },
            {
                name: 'Created', displayName: strings.CreatedAtLabel, sorting: true, isResizable: true, minWidth: 100,
                render: function (item, index, column) {
                    var date = new Date(item.Created);
                    return React.createElement(spfx_controls_react_1.FieldDateRenderer, { text: date.toLocaleDateString() });
                }
            },
            {
                name: 'LastItemModifiedDate', displayName: strings.LastChangeLabel, sorting: true, isResizable: true, minWidth: 120, linkPropertyName: 'webUrl',
                render: function (item, index, column) {
                    var date = new Date(item.LastItemModifiedDate);
                    return React.createElement(spfx_controls_react_1.FieldDateRenderer, { text: date.toLocaleString() });
                }
            },
            {
                name: 'LastItemUserModifiedDate', displayName: strings.UserChangedLabel, sorting: true, isResizable: true, minWidth: 120, linkPropertyName: 'webUrl',
                render: function (item, index, column) {
                    var date = new Date(item.LastItemUserModifiedDate);
                    return React.createElement(spfx_controls_react_1.FieldDateRenderer, { text: date.toLocaleString() });
                }
            },
            {
                name: 'LastItemDeletedDate', displayName: strings.LastDeletionLabel, sorting: true, isResizable: true, minWidth: 100,
                render: function (item, index, column) {
                    var date = new Date(item.LastItemDeletedDate);
                    return React.createElement(spfx_controls_react_1.FieldDateRenderer, { text: date.toLocaleString() });
                }
            },
            { name: 'ItemCount', displayName: 'Items', sorting: true, isResizable: true, minWidth: 120 },
            { name: 'FoundItems', displayName: strings.FoundLabel, sorting: true, isResizable: true, minWidth: 120,
                render: function (item, index, column) {
                    var _a;
                    var entry = _this.GetLibraryEntryByIndex(item.Id);
                    if (typeof entry.FoundItems !== "undefined" && entry.FoundItems !== null) {
                        return React.createElement(spfx_controls_react_1.FieldTextRenderer, { text: "".concat(strings.FoundLabel, ": ").concat((_a = entry.FoundItems) === null || _a === void 0 ? void 0 : _a.length) });
                    }
                    else
                        return React.createElement(spfx_controls_react_1.FieldTextRenderer, { text: strings.StartQueryForResults });
                }
            },
            { name: 'Description', displayName: 'Description', sorting: true, isResizable: true, minWidth: 100 }
        ];
        _this.viewFieldsPage = [
            { name: 'title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 120 },
            { name: 'name', displayName: 'Name', sorting: true, isResizable: true, minWidth: 100 },
            { name: 'webUrl', displayName: 'URL', sorting: false, isResizable: true, minWidth: 200 },
            { name: 'Links', displayName: 'Links', sorting: false, isResizable: true, minWidth: 200,
                render: function (item, index, column) {
                    var entry = _this.state.pageResults.filter(function (x) { return x.pageID === item.id; })[0];
                    if (typeof entry === "undefined" || typeof entry.Links === "undefined") {
                        return React.createElement(React.Fragment, null,
                            React.createElement(react_icons_1.CheckmarkCircleHintRegular, null));
                    }
                    if (entry.Links.filter(function (x) { return x.IsBroken; }).length > 0) {
                        return (React.createElement(React.Fragment, null,
                            React.createElement(react_icons_1.WarningColor, null),
                            "\u00A0",
                            React.createElement("span", null, strings.FoundLinksCount.replace('{0}', entry.Links.length.toString()).replace('{1}', entry.Links.filter(function (x) { return x.IsBroken; }).length.toString()))));
                    }
                    return React.createElement(React.Fragment, null,
                        React.createElement(react_icons_1.CheckmarkCircleColor, null),
                        "\u00A0",
                        React.createElement("span", null, strings.FoundLinksCount.replace('{0}', entry.Links.length.toString()).replace('{1}', entry.Links.filter(function (x) { return x.IsBroken; }).length.toString())));
                }
            }
        ];
        _this.viewFieldsPermissions = [
            { name: 'displayName', displayName: strings.PrincipalNameLabel, sorting: true, isResizable: true, minWidth: 180 },
            { name: 'isGroup', displayName: strings.PrincipalTypeLabel, sorting: true, isResizable: true, minWidth: 100,
                render: function (item) { return (React.createElement("span", { style: { display: 'flex', alignItems: 'center', gap: 4 } },
                    item.isGroup ? React.createElement(react_icons_1.PeopleTeam16Regular, null) : React.createElement(react_icons_1.Person16Regular, null),
                    React.createElement("span", null, item.isGroup ? strings.GroupLabel : strings.UserLabel))); }
            },
            { name: 'loginName', displayName: strings.LoginNameLabel, sorting: false, isResizable: true, minWidth: 220 },
            { name: 'roles', displayName: strings.RolesLabel, sorting: false, isResizable: true, minWidth: 200,
                render: function (item) { return React.createElement(spfx_controls_react_1.FieldTextRenderer, { text: (item.roles || []).join(', ') }); }
            }
        ];
        _this.viewFieldsGroupMembers = [
            { name: 'displayName', displayName: strings.PrincipalNameLabel, sorting: true, isResizable: true, minWidth: 180 },
            { name: 'email', displayName: strings.EmailLabel, sorting: true, isResizable: true, minWidth: 220 },
            { name: 'loginName', displayName: strings.LoginNameLabel, sorting: false, isResizable: true, minWidth: 220 }
        ];
        _this.handleTreeOpenChange = function (_event, data) {
            var openKeys = data.openItems;
            _this.setState({ openTreeNodeKeys: openKeys });
            if (data.open) {
                var node = _this.findTreeNode(_this.state.permissionGroupTree, String(data.value));
                if (node && node.children === undefined && !node.isLoadingChildren) {
                    void _this.loadNestedGroups(node);
                }
            }
        };
        _this.onDropdDownSelectionChanged = function (event, data) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var dataManager, pages, siteInfo, libraries;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        dataManager = new GraphDataManager_1.default(this.props.msGraphClientFactory, this.props.spHTTPClient);
                        this.setState({ isFilteringLibraries: true });
                        return [4 /*yield*/, dataManager.GetPages4Site(data.optionValue)];
                    case 1:
                        pages = _a.sent();
                        this.setState({
                            selectedTabValue: this.state.selectedTabValue === null ? "tab1" : this.state.selectedTabValue,
                            pageEntries: pages,
                            selectedSiteId: data.optionValue
                        });
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, , 4, 5]);
                        siteInfo = this.state.SelectedSites.filter(function (x) { return x.id === data.optionValue; })[0];
                        return [4 /*yield*/, dataManager.GetAllLists(siteInfo.url, this.state.chkShowLists, this.state.chkShowLibaries)];
                    case 3:
                        libraries = _a.sent();
                        console.log("All lists", libraries);
                        this.setState({
                            libraryEntries: libraries
                        });
                        return [3 /*break*/, 5];
                    case 4:
                        this.setState({ isFilteringLibraries: false });
                        return [7 /*endfinally*/];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        _this.onListSelectionChanged = function (items) {
            var selected = (items && items.length > 0) ? items[0] : null;
            _this.setState({ selectedPage: selected });
        };
        _this.onLibrarySelectionChanged = function (items) {
            var selected = (items && items.length > 0) ? items[0] : null;
            if (selected !== null)
                _this.setState({ selectedLibrary: _this.GetLibraryEntryByIndex(selected.Id) });
            else
                _this.setState({ selectedLibrary: null });
        };
        _this.onTabSelect = function (event, data) {
            _this.setState({ selectedTabValue: data.value });
        };
        _this.onFoundItemSelectionChanged = function (items) {
            var selected = (items && items.length > 0) ? items[0] : null;
            _this.setState({ selectedFoundItem: selected });
        };
        _this.onShowPermissionsClick = function () { return tslib_1.__awaiter(_this, void 0, void 0, function () {
            var site;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.state.selectedFoundItem || !this.state.selectedLibrary) {
                            console.warn('No item selected or no library selected');
                            return [2 /*return*/];
                        }
                        site = this.GetSelectedSite();
                        return [4 /*yield*/, this.GetPermission4SelectedItem(site, this.state.selectedLibrary.Id, this.state.selectedFoundItem.Id)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        }); };
        _this.state = {
            dateStartDate: new Date(),
            pageResults: [],
            SelectedSites: _this.tempSelectedSites,
            libraryEntries: [],
            selectedSiteId: null,
            isReportOpen: false,
            selectedPage: null,
            isLibraryReportOpen: false,
            selectedLibrary: null,
            selectedTabValue: null,
            pageEntries: [],
            chkShowLibaries: true,
            chkShowLists: true,
            selectedFoundItem: null,
            isQueryingLibraries: false,
            isFilteringLibraries: false,
            isProcessingBrokenLinks: false,
            expandedContentSections: new Set(),
            showOnlyBrokenLinks: false,
            isPagePermissionsOpen: false,
            pagePermissions: [],
            isLoadingPagePermissions: false,
            pagePermissionsError: null,
            permissionGroupTree: [],
            openTreeNodeKeys: new Set(),
            selectedTreeNodeKey: 'root',
            groupMemberCache: new Map(),
            isLoadingGroupMembers: false,
            groupMembersError: null,
            currentArtefact: null,
            permissionsSubjectTitle: '',
            isCheckingPrincipalAccess: false,
            principalAccessResult: null,
            principalAccessError: null
        };
        _this.dataManager = new GraphDataManager_1.default(_this.props.msGraphClientFactory, _this.props.spHTTPClient);
        _this.permissionsManager = new PermissionsManager_1.default(_this.props.msGraphClientFactory, _this.props.spHTTPClient);
        return _this;
    }
    ContentHealthManager.prototype.GetLibraryEntryByIndex = function (index) {
        return this.state.libraryEntries.filter(function (x) { return x.Id === index; })[0];
    };
    /**https://storybooks.fluentui.dev/react/?path=/docs/components-tablist--docs*/
    ContentHealthManager.prototype.render = function () {
        var _this = this;
        return (React.createElement("section", { className: ContentHealthManager_module_scss_1.default.contentHealthManager },
            this.state.SelectedSites.length === 0 && (React.createElement("div", { className: ContentHealthManager_module_scss_1.default.summarySection },
                React.createElement("div", { className: ContentHealthManager_module_scss_1.default.summaryDescription },
                    React.createElement(react_icons_1.Search24Regular, { className: ContentHealthManager_module_scss_1.default.summaryIcon }),
                    React.createElement("div", null,
                        React.createElement("h3", null, strings.ContentHealthManagerTitle),
                        React.createElement("p", null,
                            React.createElement(react_icons_1.DataTrending24Regular, { className: ContentHealthManager_module_scss_1.default.inlineIcon }),
                            strings.ContentHealthManagerDescription))),
                React.createElement("div", { className: ContentHealthManager_module_scss_1.default.instructionsSection },
                    React.createElement("h4", null,
                        React.createElement(react_icons_1.List24Regular, { className: ContentHealthManager_module_scss_1.default.inlineIcon }),
                        strings.HowToUseHeading),
                    React.createElement("ol", { className: ContentHealthManager_module_scss_1.default.stepList },
                        React.createElement("li", null,
                            React.createElement("strong", null, strings.FirstSelectSites.split(' - ')[0]),
                            " - ",
                            strings.FirstSelectSites.split(' - ')[1]),
                        React.createElement("li", null,
                            React.createElement("strong", null, strings.SecondSelectSingleSite.split(' - ')[0]),
                            " - ",
                            strings.SecondSelectSingleSite.split(' - ')[1]),
                        React.createElement("li", null,
                            React.createElement("strong", null, strings.StartQueryToFind),
                            React.createElement("ul", { className: ContentHealthManager_module_scss_1.default.subList },
                                React.createElement("li", null,
                                    React.createElement(react_icons_1.Link24Regular, { className: ContentHealthManager_module_scss_1.default.inlineIcon }),
                                    strings.BrokenLinksInPages),
                                React.createElement("li", null,
                                    React.createElement(react_icons_1.Clock24Regular, { className: ContentHealthManager_module_scss_1.default.inlineIcon }),
                                    strings.OldContentForDate),
                                React.createElement("li", null,
                                    React.createElement(react_icons_1.LockClosed24Regular, { className: ContentHealthManager_module_scss_1.default.inlineIcon }),
                                    strings.CheckedOutContentItems))))))),
            React.createElement("div", { className: ContentHealthManager_module_scss_1.default.row },
                React.createElement("div", { className: ContentHealthManager_module_scss_1.default['col-sm12'] },
                    this.state.SelectedSites.length === 0 && React.createElement("p", { className: ContentHealthManager_module_scss_1.default.infoMessage },
                        React.createElement(react_icons_1.QuestionCircleColor, null),
                        strings.SelectFirstAllSites),
                    React.createElement(react_components_1.Field, { label: strings.SelectSitesLabel },
                        React.createElement(SitePicker_1.SitePicker, { context: this.props.wpContext, mode: 'site', selectedSites: this.tempSelectedSites, allowSearch: true, multiSelect: true, className: ContentHealthManager_module_scss_1.default.sitePicker, trimDuplicates: true, onChange: function (sites) {
                                console.log(sites);
                                _this.setState({ SelectedSites: sites });
                            }, placeholder: strings.SelectAllSitesPlaceholder, searchPlaceholder: strings.FilterSitesPlaceholder }))),
                React.createElement("div", { className: ContentHealthManager_module_scss_1.default['col-sm12'] },
                    this.state.SelectedSites.length > 0 && this.state.selectedSiteId === null && React.createElement("div", null,
                        React.createElement("p", { className: ContentHealthManager_module_scss_1.default.infoMessage },
                            React.createElement(react_icons_1.QuestionCircleColor, null),
                            strings.ToContinueSelectSite)),
                    this.state.SelectedSites.length > 0 &&
                        React.createElement(react_components_1.Field, { label: strings.ChooseSiteLabel },
                            React.createElement(react_components_1.Dropdown, { id: 'ddCurrentSite', inlinePopup: true, onOptionSelect: this.onDropdDownSelectionChanged, placeholder: strings.SelectSitePlaceholder }, this.state.SelectedSites.map(function (entry) { return (React.createElement(react_components_1.Option, { value: entry.id, key: entry.webId }, entry.title)); }))))),
            this.state.selectedSiteId && React.createElement(React.Fragment, null,
                React.createElement("p", { className: ContentHealthManager_module_scss_1.default.infoMessage },
                    React.createElement(react_icons_1.FlagPrideIntersexInclusiveProgressFilled, null),
                    strings.ResultsForSite,
                    React.createElement("a", { href: this.GetSelectedSite().url, target: '_blank', rel: 'noreferrer' },
                        React.createElement("strong", null, this.GetSelectedSite().title))),
                React.createElement(react_components_1.TabList, { selectedValue: this.state.selectedTabValue, onTabSelect: this.onTabSelect },
                    React.createElement(react_components_1.Tab, { value: "tab1" }, strings.BrokenLinksAnalysisTab),
                    React.createElement(react_components_1.Tab, { value: "tab2" }, strings.LibraryAnalysisTab)),
                " "),
            this.state.selectedTabValue === 'tab2' && (React.createElement("div", { id: "Register1", className: ContentHealthManager_module_scss_1.default.row },
                React.createElement("div", { className: ContentHealthManager_module_scss_1.default.row },
                    React.createElement("div", { className: ContentHealthManager_module_scss_1.default['col-sm12'] },
                        React.createElement("div", { className: ContentHealthManager_module_scss_1.default.noteBox },
                            React.createElement(react_icons_1.Info24Regular, { className: ContentHealthManager_module_scss_1.default.noteBoxIcon }),
                            React.createElement("span", null, strings.SelectDateHint)))),
                React.createElement("div", { className: "".concat(ContentHealthManager_module_scss_1.default.row, " ").concat(ContentHealthManager_module_scss_1.default.libraryCommands) },
                    React.createElement("div", { className: ContentHealthManager_module_scss_1.default['col-sm5'] },
                        React.createElement(react_components_1.Field, { label: strings.SelectDateLabel, orientation: "horizontal" },
                            React.createElement(react_1.DatePicker, { value: this.state.dateStartDate, minDate: new Date(2000, 0, 1), maxDate: new Date(), placeholder: strings.SelectQueryDatePlaceholder, onSelectDate: function (selectedDate) { return _this.setState({ dateStartDate: selectedDate }); } }))),
                    React.createElement("div", { className: "".concat(ContentHealthManager_module_scss_1.default['col-sm7'], " ").concat(ContentHealthManager_module_scss_1.default.libraryCommandsLeft) },
                        React.createElement(react_components_1.Tooltip, { content: this.state.selectedLibrary ? strings.TooltipQueryLibrary : strings.TooltipQueryAllLibraries, relationship: "label" },
                            React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.DatabaseSearch24Regular, null), onClick: function () { return _this.StartQueryLstAndLibraries(); }, disabled: this.state.isQueryingLibraries },
                                !this.state.selectedLibrary && React.createElement("span", null, strings.QueryAllLibraries),
                                this.state.selectedLibrary && React.createElement("span", null, strings.QueryLibrary))),
                        this.state.isQueryingLibraries && React.createElement(react_components_1.Spinner, { size: "tiny", className: ContentHealthManager_module_scss_1.default.progressSpinner }),
                        "\u00A0",
                        React.createElement(react_components_1.Tooltip, { content: strings.TooltipCheckedOutItems, relationship: "label" },
                            React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.LockClosed24Regular, null), onClick: function () { return _this.StartQueryCheckedOutItems(); } }, strings.CheckedOutItems)))),
                React.createElement("div", { className: "".concat(ContentHealthManager_module_scss_1.default.row, " ").concat(ContentHealthManager_module_scss_1.default.libraryCommands) },
                    React.createElement("div", { className: ContentHealthManager_module_scss_1.default['col-sm4'] },
                        React.createElement(react_components_1.Tooltip, { content: strings.TooltipOpenLibraryDetails, relationship: "label" },
                            React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.Open24Regular, null), onClick: function () { return _this.ShowLibraryReport(); }, disabled: !this.state.selectedLibrary }, strings.OpenDetails))),
                    React.createElement("div", { className: "".concat(ContentHealthManager_module_scss_1.default['col-sm8'], " ").concat(ContentHealthManager_module_scss_1.default.checkboxContainer) },
                        React.createElement(react_1.Checkbox, { checked: this.state.chkShowLibaries, disabled: this.state.isFilteringLibraries, onChange: function (ev, checked) {
                                void _this.UpdateLibraryFilter(checked || false, _this.state.chkShowLists);
                            }, label: strings.LibrariesCheckbox }),
                        React.createElement(react_1.Checkbox, { checked: this.state.chkShowLists, disabled: this.state.isFilteringLibraries, onChange: function (ev, checked) {
                                void _this.UpdateLibraryFilter(_this.state.chkShowLibaries, checked || false);
                            }, label: strings.ListsCheckbox }),
                        this.state.isFilteringLibraries && React.createElement(react_components_1.Spinner, { size: "tiny", className: ContentHealthManager_module_scss_1.default.progressSpinner }))),
                React.createElement(ListView_1.ListView, { items: this.state.libraryEntries, viewFields: this.viewFieldsLibs, compact: true, selectionMode: react_1.SelectionMode.single, selection: this.onLibrarySelectionChanged }))),
            this.state.selectedTabValue === 'tab1' && (React.createElement("div", { id: "Register2", className: ContentHealthManager_module_scss_1.default.row },
                React.createElement("div", { className: "".concat(ContentHealthManager_module_scss_1.default.row, " ").concat(ContentHealthManager_module_scss_1.default.libraryCommands) },
                    React.createElement("div", { className: "".concat(ContentHealthManager_module_scss_1.default['col-sm12'], " ").concat(ContentHealthManager_module_scss_1.default.libraryCommandsLeft) },
                        React.createElement(react_components_1.Tooltip, { content: this.state.selectedPage ? strings.TooltipProcessPage : strings.TooltipFindBrokenLinks, relationship: "label" },
                            React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.Link24Regular, null), onClick: function () { return _this.StartBrokenLinkProcess(); }, disabled: this.state.isProcessingBrokenLinks },
                                !this.state.selectedPage && React.createElement("span", null, strings.FindBrokenLinks),
                                this.state.selectedPage && React.createElement("span", null, strings.ProcessPage))),
                        this.state.isProcessingBrokenLinks && React.createElement(react_components_1.Spinner, { size: "tiny", className: ContentHealthManager_module_scss_1.default.progressSpinner }),
                        "\u00A0",
                        React.createElement(react_components_1.Tooltip, { content: strings.TooltipOpenPageDetails, relationship: "label" },
                            React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.Open24Regular, null), onClick: function () { return _this.ShowPageReport(); }, disabled: !this.state.selectedPage }, strings.OpenDetails)),
                        "\u00A0",
                        React.createElement(react_components_1.Tooltip, { content: this.state.selectedPage ? strings.TooltipShowPermissions : strings.TooltipShowLibraryPermissions, relationship: "label" },
                            React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.KeyMultiple24Regular, null), onClick: function () { return _this.ShowPagePermissions(); } }, strings.PermissionsButtonLabel)))),
                React.createElement(ListView_1.ListView, { items: this.state.pageEntries, viewFields: this.viewFieldsPage, compact: true, selectionMode: react_1.SelectionMode.single, selection: this.onListSelectionChanged }))),
            React.createElement(react_components_1.Dialog, { open: !!this.state.isReportOpen, onOpenChange: function (_, data) { return _this.setState({ isReportOpen: !!data.open }); }, modalType: 'alert' },
                React.createElement(react_components_1.DialogSurface, null,
                    React.createElement(react_components_1.DialogBody, null,
                        React.createElement(react_components_1.DialogTitle, null, strings.PageReportTitle),
                        React.createElement(react_components_1.DialogContent, { style: { padding: 12 } }, this.state.selectedPage ? (React.createElement("div", null,
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.TitleLabel),
                                " ",
                                this.state.selectedPage.title || this.state.selectedPage.name),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.UrlLabel),
                                " ",
                                React.createElement("a", { href: this.state.selectedPage.webUrl, target: '_blank', rel: 'noreferrer' }, this.state.selectedPage.webUrl)),
                            (function () {
                                var entry = _this.state.pageResults.filter(function (x) { return x.pageID === _this.state.selectedPage.id; })[0];
                                if (entry) {
                                    return (React.createElement("div", { style: { marginTop: 8 } },
                                        React.createElement("div", null,
                                            React.createElement("strong", null, strings.TotalLinksLabel),
                                            " ",
                                            entry.Links.length),
                                        React.createElement("div", null,
                                            React.createElement("strong", null, strings.BrokenLinksLabel),
                                            " ",
                                            entry.Links.filter(function (l) { return l.IsBroken; }).length),
                                        React.createElement("div", { style: { marginTop: 12 } },
                                            React.createElement("div", { style: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' } },
                                                React.createElement("div", null,
                                                    React.createElement("strong", null, strings.AllLinksLabel)),
                                                React.createElement(react_1.Toggle, { checked: _this.state.showOnlyBrokenLinks, onChange: function (ev, checked) {
                                                        _this.setState({ showOnlyBrokenLinks: checked || false });
                                                    }, label: strings.ShowOnlyBrokenLinks, inlineLabel: true })),
                                            React.createElement("div", { style: { maxHeight: '300px', overflowY: 'auto', marginTop: 8, border: '1px solid #ccc', padding: 8 } },
                                                (function () {
                                                    var filteredLinks = _this.state.showOnlyBrokenLinks
                                                        ? entry.Links.filter(function (l) { return l.IsBroken; })
                                                        : entry.Links;
                                                    return filteredLinks.length > 0 ? (filteredLinks.map(function (link, index) { return (React.createElement("div", { key: index, style: {
                                                            padding: '8px',
                                                            marginBottom: '4px',
                                                            border: '1px solid #e0e0e0',
                                                            borderRadius: '4px',
                                                            backgroundColor: link.IsBroken ? '#ffebee' : '#f5f5f5'
                                                        } },
                                                        React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: '8px' } },
                                                            React.createElement("span", { style: {
                                                                    color: link.IsBroken ? '#d32f2f' : '#2e7d32',
                                                                    fontWeight: 'bold',
                                                                    fontSize: '12px'
                                                                } }, link.IsBroken ? '❌ BROKEN' : '✅ OK')),
                                                        React.createElement("div", { style: { marginTop: '4px' } },
                                                            React.createElement("div", null,
                                                                React.createElement("strong", null, strings.TitleLabel),
                                                                " ",
                                                                link.title || strings.NoTitle),
                                                            React.createElement("div", null,
                                                                React.createElement("strong", null, strings.UrlLabel),
                                                                React.createElement("a", { href: link.url, target: "_blank", rel: "noopener noreferrer", style: { marginLeft: '4px', color: '#0078d4' } }, link.title || strings.NoTitle)),
                                                            link.Content && link.Content.trim().length > 0 && (React.createElement("div", { style: { marginTop: '8px' } },
                                                                React.createElement("button", { title: strings.TooltipToggleContent, onClick: function () {
                                                                        var currentExpanded = _this.state.expandedContentSections || new Set();
                                                                        var expanded = new Set();
                                                                        currentExpanded.forEach(function (val) { return expanded.add(val); });
                                                                        if (expanded.has(link.url)) {
                                                                            expanded.delete(link.url);
                                                                        }
                                                                        else {
                                                                            expanded.add(link.url);
                                                                        }
                                                                        _this.setState({ expandedContentSections: expanded });
                                                                    }, style: {
                                                                        display: 'flex',
                                                                        alignItems: 'center',
                                                                        gap: '4px',
                                                                        background: 'none',
                                                                        border: 'none',
                                                                        cursor: 'pointer',
                                                                        color: '#0078d4',
                                                                        padding: '4px 0',
                                                                        fontSize: '14px'
                                                                    } },
                                                                    ((_this.state.expandedContentSections || new Set()).has(link.url) ? React.createElement(react_icons_1.ChevronUp24Regular, null) : React.createElement(react_icons_1.ChevronDown24Regular, null)),
                                                                    React.createElement("span", null, strings.ShowContent)),
                                                                (_this.state.expandedContentSections || new Set()).has(link.url) && (React.createElement("div", { style: {
                                                                        marginTop: '8px',
                                                                        padding: '8px',
                                                                        backgroundColor: '#f9f9f9',
                                                                        border: '1px solid #e0e0e0',
                                                                        borderRadius: '4px',
                                                                        maxHeight: '300px',
                                                                        overflowY: 'auto'
                                                                    }, dangerouslySetInnerHTML: { __html: link.Content } }))))))); })) : null;
                                                })(),
                                                (function () {
                                                    var filteredLinks = _this.state.showOnlyBrokenLinks
                                                        ? entry.Links.filter(function (l) { return l.IsBroken; })
                                                        : entry.Links;
                                                    return filteredLinks.length === 0 ? (React.createElement("div", { style: { padding: '8px', color: '#666', fontStyle: 'italic' } }, _this.state.showOnlyBrokenLinks
                                                        ? strings.NoBrokenLinksFound
                                                        : strings.NoLinksFound)) : null;
                                                })()))));
                                }
                                return React.createElement("div", { style: { marginTop: 8 } }, strings.NoLinkAnalysisAvailable);
                            })())) : (React.createElement("div", null, strings.NoItemSelected))),
                        React.createElement(react_components_1.DialogActions, null,
                            React.createElement(react_components_1.Tooltip, { content: strings.TooltipCloseDialog, relationship: "label" },
                                React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.Dismiss24Regular, null), appearance: 'secondary', onClick: function () { return _this.setState({ isReportOpen: false }); } }, strings.CloseButton)))))),
            React.createElement(react_components_1.Dialog, { open: !!this.state.isLibraryReportOpen, onOpenChange: function (_, data) { return _this.setState({ isLibraryReportOpen: !!data.open }); }, modalType: 'alert' },
                React.createElement(react_components_1.DialogSurface, null,
                    React.createElement(react_components_1.DialogBody, null,
                        React.createElement(react_components_1.DialogTitle, null, strings.LibraryReportTitle),
                        React.createElement(react_components_1.DialogContent, { style: { padding: 12 } }, this.state.selectedLibrary ? (React.createElement("div", null,
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.TitleLabel),
                                " ",
                                this.state.selectedLibrary.Title || strings.NA),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.TemplateLabel),
                                " ",
                                ListTemplateTypes_1.ListTemplateType[this.state.selectedLibrary.BaseTemplate] || strings.NA),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.DescriptionLabel),
                                " ",
                                this.state.selectedLibrary.Description || strings.NA),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.ItemCountLabel),
                                " ",
                                this.state.selectedLibrary.ItemCount),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.CreatedLabel),
                                " ",
                                new Date(this.state.selectedLibrary.Created).toLocaleDateString()),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.LastModifiedLabel),
                                " ",
                                new Date(this.state.selectedLibrary.LastItemModifiedDate).toLocaleString()),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.LastUserModifiedLabel),
                                " ",
                                new Date(this.state.selectedLibrary.LastItemUserModifiedDate).toLocaleString()),
                            this.state.selectedLibrary.LastItemDeletedDate && (React.createElement("div", null,
                                React.createElement("strong", null, strings.LastDeletedLabel),
                                " ",
                                new Date(this.state.selectedLibrary.LastItemDeletedDate).toLocaleString())),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.EnableVersioningLabel),
                                " ",
                                this.state.selectedLibrary.EnableVersioning ? strings.Yes : strings.No),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.EnableAttachmentsLabel),
                                " ",
                                this.state.selectedLibrary.EnableAttachments ? strings.Yes : strings.No),
                            React.createElement("div", null,
                                React.createElement("strong", null, strings.EnableFolderCreationLabel),
                                " ",
                                this.state.selectedLibrary.EnableFolderCreation ? strings.Yes : strings.No),
                            React.createElement("div", { style: { marginTop: 16 } },
                                React.createElement("h4", null, strings.OverviewListEntries),
                                this.state.selectedLibrary.FoundItems && this.state.selectedLibrary.FoundItems.length > 0 ? (React.createElement("div", null,
                                    React.createElement("div", null,
                                        React.createElement("strong", null, strings.TotalItemsFound),
                                        " ",
                                        this.state.selectedLibrary.FoundItems.length),
                                    React.createElement(react_components_1.Tooltip, { content: strings.TooltipShowPermissions, relationship: "label" },
                                        React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.KeyMultiple24Regular, null), onClick: this.onShowPermissionsClick, disabled: !this.state.selectedFoundItem, appearance: "secondary", style: { marginBottom: '8px' } }, strings.ShowPermissions)),
                                    React.createElement("div", { style: { marginTop: 8, maxHeight: '300px' } },
                                        React.createElement(ListView_1.ListView, { items: this.state.selectedLibrary.FoundItems, viewFields: this.viewFieldsFoundItems, compact: true, selectionMode: react_1.SelectionMode.single, selection: this.onFoundItemSelectionChanged })))) : (React.createElement("div", { style: { padding: '16px', backgroundColor: '#f5f5f5', border: '1px solid #ddd', borderRadius: '4px', textAlign: 'center' } },
                                    React.createElement("p", { style: { margin: 0, color: '#666' } }, strings.QueryLibraryForResults)))))) : (React.createElement("div", null, strings.NoLibrarySelected))),
                        React.createElement(react_components_1.DialogActions, null,
                            React.createElement(react_components_1.Tooltip, { content: strings.TooltipCloseDialog, relationship: "label" },
                                React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.Dismiss24Regular, null), appearance: 'secondary', onClick: function () { return _this.setState({ isLibraryReportOpen: false }); } }, strings.CloseButton)))))),
            React.createElement(react_components_1.Dialog, { open: !!this.state.isPagePermissionsOpen, onOpenChange: function (_, data) { return _this.setState({ isPagePermissionsOpen: !!data.open }); }, modalType: 'alert' },
                React.createElement(react_components_1.DialogSurface, { style: { maxWidth: '95vw', width: 960 } },
                    React.createElement(react_components_1.DialogBody, null,
                        React.createElement(react_components_1.DialogTitle, null, strings.PagePermissionsTitle),
                        React.createElement(react_components_1.DialogContent, { style: { padding: 12 } },
                            React.createElement("div", null,
                                React.createElement("div", null,
                                    React.createElement("strong", null, strings.TitleLabel),
                                    " ",
                                    this.state.permissionsSubjectTitle),
                                this.state.selectedPage && (React.createElement("div", null,
                                    React.createElement("strong", null, strings.UrlLabel),
                                    " ",
                                    React.createElement("a", { href: this.state.selectedPage.webUrl, target: '_blank', rel: 'noreferrer' }, this.state.selectedPage.webUrl))),
                                this.state.currentArtefact && (React.createElement("div", { style: { marginTop: 12 } },
                                    React.createElement(PeoplePicker_1.PeoplePicker, { context: {
                                            absoluteUrl: this.state.currentArtefact.webUrl,
                                            msGraphClientFactory: this.props.msGraphClientFactory,
                                            spHttpClient: this.props.spHTTPClient
                                        }, personSelectionLimit: 1, principalTypes: [PeoplePicker_1.PrincipalType.User, PeoplePicker_1.PrincipalType.SecurityGroup, PeoplePicker_1.PrincipalType.SharePointGroup, PeoplePicker_1.PrincipalType.DistributionList], placeholder: strings.SearchUserOrGroupPlaceholder, onChange: function (items) {
                                            var item = items && items[0] ? items[0] : undefined;
                                            if (item) {
                                                void _this.checkPrincipalAccess(item);
                                            }
                                            else {
                                                _this.setState({ principalAccessResult: null, principalAccessError: null });
                                            }
                                        } }),
                                    this.state.isCheckingPrincipalAccess && React.createElement(react_components_1.Spinner, { size: "tiny", className: ContentHealthManager_module_scss_1.default.progressSpinner }),
                                    this.state.principalAccessError && (React.createElement("div", { style: { color: '#d32f2f', marginTop: 8 } }, this.state.principalAccessError)),
                                    !this.state.isCheckingPrincipalAccess && !this.state.principalAccessError && this.state.principalAccessResult && (React.createElement("div", { style: { marginTop: 8 } }, this.state.principalAccessResult.hasAccess
                                        ? strings.HasAccessLabel
                                            .replace('{0}', this.state.principalAccessResult.displayName)
                                            .replace('{1}', this.getPermissionLevelLabel(this.state.principalAccessResult.permissionInfo))
                                        : strings.NoAccessLabel.replace('{0}', this.state.principalAccessResult.displayName))))),
                                this.state.isLoadingPagePermissions && React.createElement(react_components_1.Spinner, { size: "tiny", className: ContentHealthManager_module_scss_1.default.progressSpinner }),
                                this.state.pagePermissionsError && (React.createElement("div", { style: { color: '#d32f2f', marginTop: 8 } }, this.state.pagePermissionsError)),
                                !this.state.isLoadingPagePermissions && !this.state.pagePermissionsError && this.state.pagePermissions.length === 0 && (React.createElement("div", { style: { marginTop: 8 } }, strings.NoPermissionsFound)),
                                !this.state.isLoadingPagePermissions && !this.state.pagePermissionsError && this.state.pagePermissions.length > 0 && (React.createElement(react_resizable_panels_1.PanelGroup, { direction: "horizontal", style: { height: 420, marginTop: 12 } },
                                    React.createElement(react_resizable_panels_1.Panel, { defaultSize: 30, minSize: 15, maxSize: 60 },
                                        React.createElement("div", { style: { height: '100%', overflow: 'auto', borderRight: '1px solid #e0e0e0' } },
                                            React.createElement(react_components_1.Tree, { openItems: this.state.openTreeNodeKeys, onOpenChange: this.handleTreeOpenChange, "aria-label": strings.PagePermissionsTitle },
                                                React.createElement(react_components_1.TreeItem, { itemType: "leaf", value: "root" },
                                                    React.createElement(react_components_1.TreeItemLayout, { onClick: function () { return _this.selectTreeNode('root'); }, style: this.state.selectedTreeNodeKey === 'root' ? { background: '#e0e0e0' } : undefined }, this.state.permissionsSubjectTitle)),
                                                this.state.permissionGroupTree.map(function (node) { return _this.renderGroupTreeNode(node); })))),
                                    React.createElement(react_resizable_panels_1.PanelResizeHandle, { style: { width: 6, cursor: 'col-resize', background: '#e0e0e0' } }),
                                    React.createElement(react_resizable_panels_1.Panel, null,
                                        React.createElement("div", { style: { height: '100%', overflow: 'auto', paddingLeft: 8 } }, this.state.selectedTreeNodeKey === 'root' ? (React.createElement(ListView_1.ListView, { items: this.state.pagePermissions.filter(function (p) { return !p.isGroup; }), viewFields: this.viewFieldsPermissions, compact: true, selectionMode: react_1.SelectionMode.none })) : (React.createElement(React.Fragment, null,
                                            this.state.isLoadingGroupMembers && React.createElement(react_components_1.Spinner, { size: "tiny", className: ContentHealthManager_module_scss_1.default.progressSpinner }),
                                            this.state.groupMembersError && (React.createElement("div", { style: { color: '#d32f2f', marginTop: 8 } }, this.state.groupMembersError)),
                                            !this.state.isLoadingGroupMembers && !this.state.groupMembersError && (React.createElement(ListView_1.ListView, { items: this.state.groupMemberCache.get(this.state.selectedTreeNodeKey) || [], viewFields: this.viewFieldsGroupMembers, compact: true, selectionMode: react_1.SelectionMode.none })))))))))),
                        React.createElement(react_components_1.DialogActions, null,
                            React.createElement(react_components_1.Tooltip, { content: strings.TooltipCloseDialog, relationship: "label" },
                                React.createElement(react_components_1.Button, { icon: React.createElement(react_icons_1.Dismiss24Regular, null), appearance: 'secondary', onClick: function () { return _this.setState({ isPagePermissionsOpen: false }); } }, strings.CloseButton))))))));
    };
    ContentHealthManager.prototype.componentDidMount = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                return [2 /*return*/];
            });
        });
    };
    ContentHealthManager.prototype.ShowLibraryReport = function () {
        if (!this.state.selectedLibrary) {
            return;
        }
        this.setState({ isLibraryReportOpen: true });
    };
    ContentHealthManager.prototype.ShowPageReport = function () {
        if (!this.state.selectedPage) {
            return;
        }
        this.setState({ isReportOpen: true });
    };
    ContentHealthManager.prototype.ShowPagePermissions = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var site, artefact_1, permissions, groupTree, error_1;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        site = this.GetSelectedSite();
                        if (!site) {
                            console.warn('No site selected. Please select a site first.');
                            return [2 /*return*/];
                        }
                        this.setState({
                            isPagePermissionsOpen: true,
                            isLoadingPagePermissions: true,
                            pagePermissions: [],
                            pagePermissionsError: null,
                            permissionGroupTree: [],
                            openTreeNodeKeys: new Set(),
                            selectedTreeNodeKey: 'root',
                            groupMemberCache: new Map(),
                            groupMembersError: null,
                            currentArtefact: null,
                            permissionsSubjectTitle: this.state.selectedPage ? (this.state.selectedPage.title || this.state.selectedPage.name || '') : strings.PagesLibraryLabel,
                            isCheckingPrincipalAccess: false,
                            principalAccessResult: null,
                            principalAccessError: null
                        });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 7, 8, 9]);
                        if (!this.state.selectedPage) return [3 /*break*/, 3];
                        if (!this.state.selectedPage.webUrl) {
                            throw new Error('The selected page has no URL.');
                        }
                        return [4 /*yield*/, this.permissionsManager.resolveArtefactFromFileUrl(site.url, this.state.selectedPage.webUrl)];
                    case 2:
                        artefact_1 = _a.sent();
                        return [3 /*break*/, 5];
                    case 3: return [4 /*yield*/, this.permissionsManager.resolvePagesLibraryArtefact(site.url)];
                    case 4:
                        artefact_1 = _a.sent();
                        _a.label = 5;
                    case 5: return [4 /*yield*/, this.permissionsManager.get4ArtefactPermissions(artefact_1)];
                    case 6:
                        permissions = _a.sent();
                        groupTree = permissions.filter(function (p) { return p.isGroup; }).map(function (p) { return _this.buildGroupNode(p, artefact_1.webUrl); });
                        this.setState({ pagePermissions: permissions, permissionGroupTree: groupTree, currentArtefact: artefact_1 });
                        return [3 /*break*/, 9];
                    case 7:
                        error_1 = _a.sent();
                        console.error('Error retrieving page permissions:', error_1);
                        this.setState({ pagePermissionsError: error_1 instanceof Error ? error_1.message : String(error_1) });
                        return [3 /*break*/, 9];
                    case 8:
                        this.setState({ isLoadingPagePermissions: false });
                        return [7 /*endfinally*/];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.checkPrincipalAccess = function (item) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var report, error_2;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.state.currentArtefact) {
                            return [2 /*return*/];
                        }
                        this.setState({ isCheckingPrincipalAccess: true, principalAccessResult: null, principalAccessError: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, 4, 5]);
                        return [4 /*yield*/, this.permissionsManager.checkAccess4Principal({ id: item.id, displayName: item.text }, this.state.currentArtefact)];
                    case 2:
                        report = _a.sent();
                        this.setState({ principalAccessResult: { displayName: item.text, hasAccess: report.hasAccess, permissionInfo: report.permissionInfo } });
                        return [3 /*break*/, 5];
                    case 3:
                        error_2 = _a.sent();
                        console.error('Error checking principal access:', error_2);
                        this.setState({ principalAccessError: error_2 instanceof Error ? error_2.message : String(error_2) });
                        return [3 /*break*/, 5];
                    case 4:
                        this.setState({ isCheckingPrincipalAccess: false });
                        return [7 /*endfinally*/];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.getPermissionLevelLabel = function (info) {
        if (info.hasFullControl || info.canManagePermissions) {
            return strings.FullControlLabel;
        }
        if (info.canManageLists) {
            return strings.DesignLabel;
        }
        if (info.canEdit) {
            return strings.EditLabel;
        }
        if (info.canContribute) {
            return strings.ContributeLabel;
        }
        if (info.canView) {
            return strings.ReadLabel;
        }
        return strings.NoAccessLevelLabel;
    };
    ContentHealthManager.prototype.buildGroupNode = function (source, webUrl) {
        var groupInfo = 'webUrl' in source
            ? source
            : {
                webUrl: webUrl,
                principalId: source.principalId,
                principalType: source.principalType,
                loginName: source.loginName,
                displayName: source.displayName
            };
        var key = groupInfo.principalId !== undefined ? "id:".concat(groupInfo.principalId) : "login:".concat(groupInfo.loginName);
        return { key: key, groupInfo: groupInfo, children: undefined };
    };
    ContentHealthManager.prototype.findTreeNode = function (nodes, key) {
        for (var _i = 0, nodes_1 = nodes; _i < nodes_1.length; _i++) {
            var node = nodes_1[_i];
            if (node.key === key) {
                return node;
            }
            if (node.children) {
                var found = this.findTreeNode(node.children, key);
                if (found) {
                    return found;
                }
            }
        }
        return undefined;
    };
    ContentHealthManager.prototype.updateTreeNode = function (nodes, key, patch) {
        var _this = this;
        return nodes.map(function (node) {
            if (node.key === key) {
                return tslib_1.__assign(tslib_1.__assign({}, node), patch);
            }
            if (node.children) {
                return tslib_1.__assign(tslib_1.__assign({}, node), { children: _this.updateTreeNode(node.children, key, patch) });
            }
            return node;
        });
    };
    ContentHealthManager.prototype.loadNestedGroups = function (node) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var nestedGroups, children, error_3, message;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ permissionGroupTree: this.updateTreeNode(this.state.permissionGroupTree, node.key, { isLoadingChildren: true, loadError: null }) });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.permissionsManager.resolveNestedGroups(node.groupInfo)];
                    case 2:
                        nestedGroups = _a.sent();
                        children = nestedGroups.map(function (g) { return _this.buildGroupNode(g, node.groupInfo.webUrl); });
                        this.setState({ permissionGroupTree: this.updateTreeNode(this.state.permissionGroupTree, node.key, { children: children, isLoadingChildren: false }) });
                        return [3 /*break*/, 4];
                    case 3:
                        error_3 = _a.sent();
                        console.error('Error resolving nested groups:', error_3);
                        message = error_3 instanceof Error ? error_3.message : String(error_3);
                        this.setState({ permissionGroupTree: this.updateTreeNode(this.state.permissionGroupTree, node.key, { isLoadingChildren: false, loadError: message }) });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.selectTreeNode = function (key, groupInfo) {
        this.setState({ selectedTreeNodeKey: key });
        if (key === 'root' || !groupInfo) {
            return;
        }
        if (this.state.groupMemberCache.has(key)) {
            return;
        }
        void this.loadGroupMembers(key, groupInfo);
    };
    ContentHealthManager.prototype.loadGroupMembers = function (key, groupInfo) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var users, cache, error_4;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ isLoadingGroupMembers: true, groupMembersError: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, 4, 5]);
                        return [4 /*yield*/, this.permissionsManager.resolveUser4Group(groupInfo)];
                    case 2:
                        users = _a.sent();
                        cache = new Map(this.state.groupMemberCache);
                        cache.set(key, users);
                        this.setState({ groupMemberCache: cache });
                        return [3 /*break*/, 5];
                    case 3:
                        error_4 = _a.sent();
                        console.error('Error resolving group members:', error_4);
                        this.setState({ groupMembersError: error_4 instanceof Error ? error_4.message : String(error_4) });
                        return [3 /*break*/, 5];
                    case 4:
                        this.setState({ isLoadingGroupMembers: false });
                        return [7 /*endfinally*/];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.renderGroupTreeNode = function (node) {
        var _this = this;
        var _a, _b;
        return (React.createElement(react_components_1.TreeItem, { itemType: "branch", value: node.key, key: node.key },
            React.createElement(react_components_1.TreeItemLayout, { onClick: function () { return _this.selectTreeNode(node.key, node.groupInfo); }, style: this.state.selectedTreeNodeKey === node.key ? { background: '#e0e0e0' } : undefined, iconBefore: React.createElement(react_icons_1.PeopleTeam16Regular, null) }, node.groupInfo.displayName),
            React.createElement(react_components_1.Tree, null,
                node.isLoadingChildren && (React.createElement(react_components_1.TreeItem, { itemType: "leaf", value: "".concat(node.key, "-loading") },
                    React.createElement(react_components_1.TreeItemLayout, null,
                        React.createElement(react_components_1.Spinner, { size: "tiny" })))),
                node.loadError && (React.createElement(react_components_1.TreeItem, { itemType: "leaf", value: "".concat(node.key, "-error") },
                    React.createElement(react_components_1.TreeItemLayout, null,
                        React.createElement("span", { style: { color: '#d32f2f' } }, node.loadError)))),
                ((_a = node.children) === null || _a === void 0 ? void 0 : _a.length) === 0 && (React.createElement(react_components_1.TreeItem, { itemType: "leaf", value: "".concat(node.key, "-empty") },
                    React.createElement(react_components_1.TreeItemLayout, null, strings.NoNestedGroups))), (_b = node.children) === null || _b === void 0 ? void 0 :
                _b.map(function (child) { return _this.renderGroupTreeNode(child); }))));
    };
    ContentHealthManager.prototype.StartBrokenLinkProcess = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var pageAnalyzer, fullPageContent, resultLinks, _i, _a, pageEntry, fullPageContent, resultLinks, error_5, error_6;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!this.state.selectedSiteId) {
                            console.warn('No site selected. Please select a site first.');
                            return [2 /*return*/];
                        }
                        if (!this.state.pageEntries || this.state.pageEntries.length === 0) {
                            console.warn('No pages found for the selected site.');
                            return [2 /*return*/];
                        }
                        this.setState({ isProcessingBrokenLinks: true });
                        console.log("Starting broken link process for site: ".concat(this.state.selectedSiteId));
                        console.log("Processing ".concat(this.state.pageEntries.length, " pages..."));
                        pageAnalyzer = new PageProcessing_1.PageProcessing();
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 12, 13, 14]);
                        if (!this.state.selectedPage) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.dataManager.GetPageContent(this.state.selectedSiteId, this.state.selectedPage.id)];
                    case 2:
                        fullPageContent = _b.sent();
                        return [4 /*yield*/, pageAnalyzer.AnalyzePageContent(fullPageContent.canvasLayout)];
                    case 3:
                        resultLinks = _b.sent();
                        this.state.pageResults.push({
                            pageID: this.state.selectedPage.id,
                            Links: resultLinks
                        });
                        return [3 /*break*/, 11];
                    case 4:
                        _i = 0, _a = this.state.pageEntries;
                        _b.label = 5;
                    case 5:
                        if (!(_i < _a.length)) return [3 /*break*/, 11];
                        pageEntry = _a[_i];
                        _b.label = 6;
                    case 6:
                        _b.trys.push([6, 9, , 10]);
                        console.log("Processing page: ".concat(pageEntry.title || pageEntry.name, " (ID: ").concat(pageEntry.InProgress, ")"));
                        return [4 /*yield*/, this.dataManager.GetPageContent(this.state.selectedSiteId, pageEntry.id)];
                    case 7:
                        fullPageContent = _b.sent();
                        return [4 /*yield*/, pageAnalyzer.AnalyzePageContent(fullPageContent.canvasLayout)];
                    case 8:
                        resultLinks = _b.sent();
                        this.state.pageResults.push({
                            pageID: pageEntry.id,
                            Links: resultLinks
                        });
                        this.setState({
                            pageEntries: this.state.pageEntries
                        });
                        return [3 /*break*/, 10];
                    case 9:
                        error_5 = _b.sent();
                        console.error("Error processing page ".concat(pageEntry.title || pageEntry.name, ":"), error_5);
                        return [3 /*break*/, 10];
                    case 10:
                        _i++;
                        return [3 /*break*/, 5];
                    case 11: return [3 /*break*/, 14];
                    case 12:
                        error_6 = _b.sent();
                        console.error('Error during broken link process:', error_6);
                        return [3 /*break*/, 14];
                    case 13:
                        this.setState({ isProcessingBrokenLinks: false });
                        return [7 /*endfinally*/];
                    case 14: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.CollectItemsFromListAndLibraries = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var site, items, _i, _a, listInfo, items;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        site = this.GetSelectedSite();
                        console.log(this.state.selectedLibrary);
                        if (!this.state.selectedLibrary) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.dataManager.Query4ItemByDate(site, this.state.selectedLibrary.Id, this.state.selectedLibrary.ParentWebUrl, this.state.dateStartDate)];
                    case 1:
                        items = _b.sent();
                        this.state.selectedLibrary.FoundItems = items;
                        return [3 /*break*/, 6];
                    case 2:
                        _i = 0, _a = this.state.libraryEntries;
                        _b.label = 3;
                    case 3:
                        if (!(_i < _a.length)) return [3 /*break*/, 6];
                        listInfo = _a[_i];
                        return [4 /*yield*/, this.dataManager.Query4ItemByDate(site, listInfo.Id, listInfo.ParentWebUrl, this.state.dateStartDate)];
                    case 4:
                        items = _b.sent();
                        listInfo.FoundItems = items;
                        this.setState({
                            libraryEntries: this.state.libraryEntries
                        });
                        _b.label = 5;
                    case 5:
                        _i++;
                        return [3 /*break*/, 3];
                    case 6:
                        this.setState({
                            libraryEntries: this.state.libraryEntries
                        });
                        return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.GetCheckedOutItems = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var site, _i, _a, listInfo, items;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        site = this.GetSelectedSite();
                        _i = 0, _a = this.state.libraryEntries;
                        _b.label = 1;
                    case 1:
                        if (!(_i < _a.length)) return [3 /*break*/, 4];
                        listInfo = _a[_i];
                        return [4 /*yield*/, this.dataManager.Query4CheckedOutItems(site, listInfo.Id, listInfo.DefaultView.ServerRelativeUrl, this.state.dateStartDate)];
                    case 2:
                        items = _b.sent();
                        listInfo.FoundItems = items;
                        this.setState({
                            libraryEntries: this.state.libraryEntries
                        });
                        _b.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.StartQueryCheckedOutItems = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.GetCheckedOutItems()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.GetPermission4SelectedItem = function (site, listID, listItemID) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var permissions, error_7;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.dataManager.GetPermission4Item(site, listID, listItemID)];
                    case 1:
                        permissions = _a.sent();
                        console.log('Item permissions:', permissions);
                        return [3 /*break*/, 3];
                    case 2:
                        error_7 = _a.sent();
                        console.error('Error retrieving item permissions:', error_7);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.StartQueryLstAndLibraries = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ isQueryingLibraries: true });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, , 3, 4]);
                        return [4 /*yield*/, this.CollectItemsFromListAndLibraries()];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        this.setState({ isQueryingLibraries: false });
                        return [7 /*endfinally*/];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.UpdateLibraryFilter = function (chkShowLibaries, chkShowLists) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var libraries;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!chkShowLibaries && !chkShowLists) {
                            this.setState({ chkShowLibaries: chkShowLibaries, chkShowLists: chkShowLists, libraryEntries: [] });
                            return [2 /*return*/];
                        }
                        this.setState({ isFilteringLibraries: true, chkShowLibaries: chkShowLibaries, chkShowLists: chkShowLists });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, , 3, 4]);
                        return [4 /*yield*/, this.dataManager.GetAllLists(this.GetSelectedSite().url, chkShowLists, chkShowLibaries)];
                    case 2:
                        libraries = _a.sent();
                        this.setState({ libraryEntries: libraries });
                        return [3 /*break*/, 4];
                    case 3:
                        this.setState({ isFilteringLibraries: false });
                        return [7 /*endfinally*/];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    ContentHealthManager.prototype.GetSelectedSite = function () {
        var _this = this;
        return this.state.SelectedSites.filter(function (x) { return x.id === _this.state.selectedSiteId; })[0];
    };
    return ContentHealthManager;
}(React.Component));
exports.default = ContentHealthManager;
//# sourceMappingURL=ContentHealthManager.js.map