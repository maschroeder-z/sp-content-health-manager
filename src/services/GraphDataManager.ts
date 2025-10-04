import { ISPHttpClientOptions, MSGraphClientFactory, MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';
import type { Page } from '../models/Page';
import type { ListInformation } from '../models/REST/ListInformation';

//import * as MicrosoftGraph from "@microsoft/microsoft-graph-types-beta"; //[MicrosoftGraph.SitePage]
import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"

export class GraphDataManager {
  private readonly graphClientPromise: Promise<MSGraphClientV3>;
  private readonly spHTTPClient: SPHttpClient;

  constructor(msGraphClientFactory: MSGraphClientFactory, spHttpClient: SPHttpClient) {
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

  public async GetLibraries(siteID: string,): Promise<MicrosoftGraphBeta.List[]> {
    const client = await this.graphClientPromise;

    const response = await client
      .api(`/sites/${encodeURIComponent(siteID)}/lists`)
      .version('v1.0')
      .select(['id', 'name', 'displayName', 'webUrl', 'createdDateTime', 'lastModifiedDateTime'].join(','))
      .get();
    return response.value as MicrosoftGraphBeta.List[];
  }

  public async GetAllLists(siteUrl: string): Promise<ListInformation[]> {
    try {      
      // Ensure the siteUrl has proper format and add the REST API endpoint
      const apiUrl = `${siteUrl}/_api/web/lists?$expand=DefaultView`;
      
      const response = await fetch(apiUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json'
        },
        credentials: 'include' // Include cookies for authentication
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      
      // The SharePoint REST API returns data in a 'd' property with 'results' array
      const lists = data.d?.results || [];      
      return lists.map((list: any) => ({
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
        ParentWebUrl: list.ParentWebUrl+"/"+list.EntityTypeName
      }));
    } catch (error) {
      console.error('Error fetching lists:', error);
      throw error;
    }
  }

/**
 * Queries list items by date using SharePoint REST API
 * Endpoint: /[siteUrl]/_api/web/lists('[listID]')/GetItems(query=@v1)?@v1={'ViewXml':'<View><Query><Where><Leq><FieldRef Name=Modified/><Value Type=DateTime>[dateStart]</Value></Leq></Where></Query></View>'}&$expand=file
 */
  public async Query4ItemByDate(siteUrl: string, listID: string, defaultUrl: string, dateStart: Date): Promise<MicrosoftGraphBeta.ListItem[]> {
    if (typeof defaultUrl !== "undefined")
    {
      try {
        // Format the date for SharePoint CAML query (ISO format)
        const formattedDate = dateStart.toISOString();
        // /sites/Demo02/Freigegebene Dokumente/Forms/AllItems.aspx /sites/Demo02/FormServerTemplates/Forms/All Forms.aspx
        const temp = defaultUrl.split("/")
        temp.pop();
        temp.push("ViewForm.aspx?id=");
        defaultUrl = temp.join("/");
        
        // Construct the ViewXml query
        const viewXml = `<View><Query><Where><Leq><FieldRef Name=Modified/><Value Type=DateTime>${formattedDate}</Value></Leq></Where></Query></View>`;
        
        const options: ISPHttpClientOptions = {
          headers: {
            'odata-version':'3.0',
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json'          
          },
          body: `{'query': {          
            'ViewXml':'${viewXml}'
          }}`
        };

        // Encode the query parameter
        //const queryParam = encodeURIComponent(`{'ViewXml':'${viewXml}'}`);
        //const queryParam = `{'ViewXml':'${viewXml}'}`;
        
        // Construct the API URL
  1     //const apiUrl = `${siteUrl}/_api/web/lists('${listID}')/GetItems(query=@v1)?@v1=${queryParam}&$expand=file`;
        const apiUrl = `${siteUrl}/_api/web/lists('${listID}')/GetItems?$expand=ParentList,File`;
        
        const response = await this.spHTTPClient.post(
          apiUrl,
          SPHttpClient.configurations.v1,
          options
        );

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();
        
        // The SharePoint REST API returns data in a 'd' property with 'results' array
        const items : MicrosoftGraphBeta.ListItem[] = data.d?.results || [];  
        console.log(defaultUrl, items); 
        // todo: build url to item
        return items;
        
      } catch (error) {
        console.error('Error querying items by date:', error);
        throw error;
      }
    }
    return [];
  }

}

export default GraphDataManager;


