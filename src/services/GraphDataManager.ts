import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';
import type { Page } from '../models/Page';

export class GraphDataManager {
  private readonly graphClientPromise: Promise<MSGraphClientV3>;

  constructor(msGraphClientFactory: MSGraphClientFactory) {
    this.graphClientPromise = msGraphClientFactory.getClient('3');
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
  public async GetPageContent(siteID: string, pageID:string): Promise<Page> {
    const client = await this.graphClientPromise;

    const response = await client
      .api(`/sites/${encodeURIComponent(siteID)}/pages/${pageID}/microsoft.graph.sitePage?$expand=canvasLayout`)
      .version('v1.0')
      .select(['id', 'name', 'title', 'webUrl', 'createdDateTime', 'lastModifiedDateTime'].join(','))
      .get();
    return response as Page;
  }

  public async GetPages4Site(siteID: string): Promise<Page[]> {
    const client = await this.graphClientPromise;

    const response = await client
      .api(`/sites/${encodeURIComponent(siteID)}/pages/microsoft.graph.sitePage`)
      .version('v1.0')
      .select(['id', 'name', 'title', 'webUrl', 'createdDateTime', 'lastModifiedDateTime'].join(','))
      .get();

    const items: Page[] = (response?.value || []).map((p: any) => ({
      id: p.id,
      name: p.name,
      title: p.title,
      webUrl: p.webUrl,
      createdDateTime: p.createdDateTime,
      lastModifiedDateTime: p.lastModifiedDateTime,      
      InProgress: false
    }));
    return items;
  }
}

export default GraphDataManager;


