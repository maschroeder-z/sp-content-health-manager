import { CanvasStructure } from "../models/CanvasStructure";
import { LinkInfo } from "../models/LinkInfo"; 
//import * as MicrosoftGraph from "@microsoft/microsoft-graph-types-beta"; [MicrosoftGraph.SitePage]
//import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
//import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"

export class PageProcessing
{
    public async AnalyzePageContent(canvas: CanvasStructure) : Promise<LinkInfo[]|null>
    {        
        if (!canvas || !Array.isArray(canvas.horizontalSections)) {
            return null;
        }

        let links :  LinkInfo[] = [];
        for (const section of canvas.horizontalSections || []) {
            for (const column of section.columns || []) {
                for (const webpart of column.webparts || []) {
                    if (webpart.innerHtml && typeof webpart.innerHtml === 'string' && webpart.innerHtml.trim().length > 0) {
                        links=links.concat(this.ExtractLinksFromContent(webpart.innerHtml));
                    }
                    if (typeof webpart.data !== "undefined" && webpart.data !== null)
                    {   
                        const propTitle = (webpart.data.properties as any).Titel !== "undefined" ? (webpart.data.properties as any).Titel : ((webpart.data.properties as any).Title !== "undefined" ? (webpart.data.properties as any).Title : "-");
                        console.log(webpart.data,webpart.data?.serverProcessedContent.links);
                        for (const link of webpart.data?.serverProcessedContent.links!) {                            
                            links.push({
                                IsBroken: false,
                                title: `${webpart.data.title} / (${propTitle}) -> ${link.key}`,
                                url: link.value
                            });                            
                        }
                    }
                }
            }
        }
        await this.CheckLinks(links);
        return links;
    }

    private async CheckLinks(links: LinkInfo[]):Promise<void>
    {
        if (!links || links.length === 0) {
            return;
        }

        //const timeoutMs = 8000;

        const check = async (link: LinkInfo): Promise<void> => {
            const url = link.url;
            if (!url) { return; }

            const doFetch = async (method: 'HEAD') => {
                //const controller = new AbortController();
                //const timeout = setTimeout(() => controller.abort(), timeoutMs);                
                try {
                    const resp = await fetch(url, { method, mode: 'no-cors' });                    
                    if (resp.status === 200 || (resp.type === "opaque" && resp.status === 0)) 
                        link.IsBroken = false;
                    else
                        link.IsBroken = true;
                } catch (e) {
                    console.log("ERROR", e);
                    link.IsBroken = true;
                } finally {

                    //clearTimeout(timeout);
                }
            };
            await doFetch('HEAD');            
        };

        // Fire checks in parallel; no need to await inside this void method
        for (const link of links) {
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            await check(link);
        }
    }

    public ExtractLinksFromContent(content: string): LinkInfo[]
    {
        if (!content || typeof content !== 'string') {
            return [];
        }

        const results: LinkInfo[] = [];

        // Capture href (different quote styles) and the inner anchor text
        const anchorRegex = /<a\b[^>]*href\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+))[^>]*>([\s\S]*?)<\/a>/gi;
        let match: RegExpExecArray | null;
        while ((match = anchorRegex.exec(content)) !== null) {
            const rawUrl = match[2] || match[3] || match[4] || '';
            const innerHtml = match[5] || '';

            const url = (rawUrl || '').trim();
            if (!url) { continue; }
            const lower = url.toLowerCase();
            if (lower.indexOf('javascript:') === 0 || lower.indexOf('mailto:') === 0) { continue; }

            const title = this.stripHtml(innerHtml).trim();
            results.push({ url, title, IsBroken: true });
        }

        return results;
    }

    private stripHtml(html: string): string
    {
        if (!html) { return ''; }
        return html.replace(/<[^>]+>/g, ' ')
                   .replace(/\s+/g, ' ');
    }
}