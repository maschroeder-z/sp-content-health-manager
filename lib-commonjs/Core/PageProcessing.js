"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PageProcessing = void 0;
var tslib_1 = require("tslib");
//import * as MicrosoftGraph from "@microsoft/microsoft-graph-types-beta"; [MicrosoftGraph.SitePage]
//import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
//import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"
var PageProcessing = /** @class */ (function () {
    function PageProcessing() {
    }
    PageProcessing.prototype.AnalyzePageContent = function (canvas) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var links, _i, _a, section, _b, _c, column, _d, _e, webpart, propTitle, _f, _g, link;
            var _h;
            return tslib_1.__generator(this, function (_j) {
                switch (_j.label) {
                    case 0:
                        if (!canvas || !Array.isArray(canvas.horizontalSections)) {
                            return [2 /*return*/, null];
                        }
                        links = [];
                        for (_i = 0, _a = canvas.horizontalSections || []; _i < _a.length; _i++) {
                            section = _a[_i];
                            for (_b = 0, _c = section.columns || []; _b < _c.length; _b++) {
                                column = _c[_b];
                                for (_d = 0, _e = column.webparts || []; _d < _e.length; _d++) {
                                    webpart = _e[_d];
                                    if (webpart.innerHtml && typeof webpart.innerHtml === 'string' && webpart.innerHtml.trim().length > 0) {
                                        links = links.concat(this.ExtractLinksFromContent(webpart.innerHtml));
                                    }
                                    if (typeof webpart.data !== "undefined" && webpart.data !== null) {
                                        propTitle = webpart.data.properties.Titel !== "undefined" ? webpart.data.properties.Titel : (webpart.data.properties.Title !== "undefined" ? webpart.data.properties.Title : "-");
                                        if ((_h = webpart.data) === null || _h === void 0 ? void 0 : _h.serverProcessedContent.links) {
                                            for (_f = 0, _g = webpart.data.serverProcessedContent.links; _f < _g.length; _f++) {
                                                link = _g[_f];
                                                links.push({
                                                    IsBroken: false,
                                                    title: "".concat(webpart.data.title, " / (").concat(propTitle, ") -> ").concat(link.key),
                                                    url: link.value,
                                                    Content: ""
                                                });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        return [4 /*yield*/, this.CheckLinks(links)];
                    case 1:
                        _j.sent();
                        return [2 /*return*/, links];
                }
            });
        });
    };
    PageProcessing.prototype.CheckLinks = function (links) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var check, _i, links_1, link;
            var _this = this;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!links || links.length === 0) {
                            return [2 /*return*/];
                        }
                        check = function (link) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
                            var url, doFetch;
                            var _this = this;
                            return tslib_1.__generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        url = link.url;
                                        if (!url) {
                                            return [2 /*return*/];
                                        }
                                        doFetch = function (method) { return tslib_1.__awaiter(_this, void 0, void 0, function () {
                                            var resp, e_1;
                                            return tslib_1.__generator(this, function (_a) {
                                                switch (_a.label) {
                                                    case 0:
                                                        _a.trys.push([0, 2, 3, 4]);
                                                        return [4 /*yield*/, fetch(url, { method: method, mode: 'no-cors' })];
                                                    case 1:
                                                        resp = _a.sent();
                                                        if (resp.status === 200 || (resp.type === "opaque" && resp.status === 0))
                                                            link.IsBroken = false;
                                                        else
                                                            link.IsBroken = true;
                                                        return [3 /*break*/, 4];
                                                    case 2:
                                                        e_1 = _a.sent();
                                                        console.log("ERROR", e_1);
                                                        link.IsBroken = true;
                                                        return [3 /*break*/, 4];
                                                    case 3: return [7 /*endfinally*/];
                                                    case 4: return [2 /*return*/];
                                                }
                                            });
                                        }); };
                                        return [4 /*yield*/, doFetch('HEAD')];
                                    case 1:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); };
                        _i = 0, links_1 = links;
                        _a.label = 1;
                    case 1:
                        if (!(_i < links_1.length)) return [3 /*break*/, 4];
                        link = links_1[_i];
                        // eslint-disable-next-line @typescript-eslint/no-floating-promises
                        return [4 /*yield*/, check(link)];
                    case 2:
                        // eslint-disable-next-line @typescript-eslint/no-floating-promises
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    PageProcessing.prototype.ExtractLinksFromContent = function (content) {
        if (!content || typeof content !== 'string') {
            return [];
        }
        var results = [];
        // Capture href (different quote styles) and the inner anchor text
        var anchorRegex = /<a\b[^>]*href\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+))[^>]*>([\s\S]*?)<\/a>/gi;
        var match;
        while ((match = anchorRegex.exec(content)) !== null) {
            var rawUrl = match[2] || match[3] || match[4] || '';
            var innerHtml = match[5] || '';
            var url = (rawUrl || '').trim();
            if (!url) {
                continue;
            }
            var lower = url.toLowerCase();
            if (lower.indexOf('javascript:') === 0 || lower.indexOf('mailto:') === 0) {
                continue;
            }
            var title = this.stripHtml(innerHtml).trim();
            results.push({ url: url, title: title, IsBroken: true, Content: content });
        }
        return results;
    };
    PageProcessing.prototype.stripHtml = function (html) {
        if (!html) {
            return '';
        }
        return html.replace(/<[^>]+>/g, ' ')
            .replace(/\s+/g, ' ');
    };
    return PageProcessing;
}());
exports.PageProcessing = PageProcessing;
//# sourceMappingURL=PageProcessing.js.map