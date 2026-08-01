"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.AppMode = void 0;
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var ReactDom = tslib_1.__importStar(require("react-dom"));
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var strings = tslib_1.__importStar(require("ContentHealthManagerWebPartStrings"));
var ContentHealthManager_1 = tslib_1.__importDefault(require("./components/ContentHealthManager"));
var react_components_1 = require("@fluentui/react-components");
var WebPartTitle_1 = require("@pnp/spfx-controls-react/lib/WebPartTitle");
require("./WebPartTitleOverrides.global.scss");
var PropertyPaneLogo_1 = tslib_1.__importDefault(require("./PropertyPaneLogo"));
var AppMode;
(function (AppMode) {
    AppMode[AppMode["SharePoint"] = 0] = "SharePoint";
    AppMode[AppMode["SharePointLocal"] = 1] = "SharePointLocal";
    AppMode[AppMode["Teams"] = 2] = "Teams";
    AppMode[AppMode["TeamsLocal"] = 3] = "TeamsLocal";
    AppMode[AppMode["Office"] = 4] = "Office";
    AppMode[AppMode["OfficeLocal"] = 5] = "OfficeLocal";
    AppMode[AppMode["Outlook"] = 6] = "Outlook";
    AppMode[AppMode["OutlookLocal"] = 7] = "OutlookLocal";
})(AppMode || (exports.AppMode = AppMode = {}));
var ContentHealthManagerWebPart = /** @class */ (function (_super) {
    tslib_1.__extends(ContentHealthManagerWebPart, _super);
    function ContentHealthManagerWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        //private _appMode: AppMode = AppMode.SharePoint;
        //private _theme: Theme = webLightTheme;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        return _this;
    }
    ContentHealthManagerWebPart.prototype.render = function () {
        var _this = this;
        var element = React.createElement(ContentHealthManager_1.default, {
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName,
            msGraphClientFactory: this.context.msGraphClientFactory,
            wpContext: this.context,
            spHTTPClient: this.context.spHttpClient
        });
        var titleElement = React.createElement(WebPartTitle_1.WebPartTitle, {
            displayMode: this.displayMode,
            title: this.properties.title || strings.ContentHealthManagerTitle,
            updateProperty: function (value) {
                _this.properties.title = value;
            }
        });
        var fluentElement = React.createElement(react_components_1.FluentProvider, {
            theme: this._isDarkTheme ? react_components_1.teamsDarkTheme : react_components_1.teamsLightTheme
        }, titleElement, element);
        var temp = React.createElement(react_components_1.IdPrefixProvider, { value: "msz" }, fluentElement);
        ReactDom.render(temp, this.domElement);
    };
    ContentHealthManagerWebPart.prototype.onInit = function () {
        var _this = this;
        return this._getEnvironmentMessage().then(function (message) {
            _this._environmentMessage = message;
        });
    };
    ContentHealthManagerWebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(function (context) {
                var environmentMessage = '';
                switch (context.app.host.name) {
                    case 'Office': // running in Office
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
                        break;
                    case 'Outlook': // running in Outlook
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
                        break;
                    case 'Teams': // running in Teams
                    case 'TeamsModern':
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
                        break;
                    default:
                        environmentMessage = strings.UnknownEnvironment;
                }
                return environmentMessage;
            });
        }
        return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
    };
    ContentHealthManagerWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors, palette = currentTheme.palette;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
        if (palette) {
            this.domElement.style.setProperty('--themePrimary', palette.themePrimary || null);
            this.domElement.style.setProperty('--themeLighterAlt', palette.themeLighterAlt || null);
            this.domElement.style.setProperty('--themeLighter', palette.themeLighter || null);
            this.domElement.style.setProperty('--neutralLighter', palette.neutralLighter || null);
            this.domElement.style.setProperty('--neutralLight', palette.neutralLight || null);
            this.domElement.style.setProperty('--neutralSecondary', palette.neutralSecondary || null);
        }
    };
    ContentHealthManagerWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(ContentHealthManagerWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    ContentHealthManagerWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: "",
                            groupFields: [
                                new PropertyPaneLogo_1.default()
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return ContentHealthManagerWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = ContentHealthManagerWebPart;
//# sourceMappingURL=ContentHealthManagerWebPart.js.map