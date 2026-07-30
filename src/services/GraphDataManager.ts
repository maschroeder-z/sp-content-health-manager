import { ISPHttpClientOptions, MSGraphClientFactory, MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';
import type { Page } from '../models/Page';
import type { ListInformation } from '../models/REST/ListInformation';

//import * as MicrosoftGraph from "@microsoft/microsoft-graph-types-beta"; //[MicrosoftGraph.SitePage]
import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"
import { Site } from '../models/Site';

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
  public async GetPageContent(siteID: string, pageID: string): Promise<Page> {
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

  public async GetAllLists(siteUrl: string, incLists: boolean, incLibraries: boolean): Promise<ListInformation[]> {
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
      const lists = data.d?.results.filter((x: any) => (x.BaseType === 0 && incLists) || (x.BaseTemplate === 101 && x.BaseType === 1 && incLibraries)) || [];
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
        ParentWebUrl: list.ParentWebUrl + "/" + list.EntityTypeName
      }));
    } catch (error) {
      console.error('Error fetching lists:', error);
      throw error;
    }
  }

  /**
   * Queries checked-out items using classic SharePoint REST instead of Graph.
   * Graph's /items endpoint rejects $filter on person-field-derived names like
   * CheckoutUserLookupId ("A provided field name is not recognized"). Classic REST also
   * rejects $select/$expand=CheckoutUser ("field or property does not exist") - the
   * checkout person field's navigation property is actually named CheckoutUserId (the "Id"
   * suffix is part of the nav property name here, not just the lookup id column), and
   * unlike File/CheckedOutByUser it filters and expands directly on the /items collection.
   */
  public async Query4CheckedOutItems(site: Site, listID: string, defaultUrl: string, dateStart: Date): Promise<MicrosoftGraphBeta.ListItem[]> {
    defaultUrl = site.url + "/_layouts/15/listform.aspx?PageType=4&ListId=";
    try {
      // Title isn't selected here: some libraries have the Title column disabled, which
      // makes classic REST reject $select=Title outright. FileLeafRef (the file name) is
      // always present and is used as the display title instead, same as the original
      // Graph-based query did.
      const apiUrl = `${site.url}/_api/web/lists('${listID}')/items` +
        `?$select=Id,FileLeafRef,Created,Modified,ContentTypeId,CheckoutUser/Title,CheckoutUser/EMail` +
        `&$expand=CheckoutUser,ContentType` +
        `&$filter=CheckoutUserId ne null`;

      const response = await this.spHTTPClient.get(
        apiUrl,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      console.log("check-out items:", data);
      const items: MicrosoftGraphBeta.ListItem[] = (data.value || []).map((item: any) => ({
        ...item,
        Title: item.FileLeafRef,
        CheckedOutBy: item.CheckoutUser?.Title || null, // EMail
        webUrl: `${defaultUrl}${listID}&id=${item.Id}`
      }));

      return items;
    } catch (error) {
      console.error('Error querying checked-out items:', error);
      throw error;
    }
  }

  public async Query4ItemByDate(site: Site, listID: string, defaultUrl: string, dateStart: Date): Promise<MicrosoftGraphBeta.ListItem[]> {
    try {
      const client = await this.graphClientPromise;

      // Format the date for Graph API filter (ISO format)
      const formattedDate = dateStart.toISOString();

      const urlToDetails: string[] = defaultUrl.split("/");
      urlToDetails.pop();
      const urlDispForm: string = urlToDetails.join("/") + "/_layouts/15/listform.aspx?PageType=4";

      // Query for items not modified after the given date using Microsoft Graph API
      const response = await client
        .api(`/sites/${encodeURIComponent(site.id)}/lists/${listID}/items`)
        .version('v1.0')
        .filter(`fields/Modified le '${formattedDate}'`)
        .expand('fields')
        .select(['id', 'fields'])
        .get();

      const items: MicrosoftGraphBeta.ListItem[] = (response?.value || []).map((item: any) => ({
        Id: item.id,
        Title: item.fields.Title || item.fields.FileLeafRef,
        Created: item.fields.Created,
        Modified: item.fields.Modified,
        ContentTypeId: item.fields.ContentTypeId,
        ...item.fields,
        webUrl: `${urlDispForm}&id=${item.id}&listid=${listID}`
      }));

      return items;
    } catch (error) {
      console.error('Error querying items by date:', error);
      throw error;
    }
  }

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
  }

  /**
 * Queries list items by date using SharePoint REST API
 * Endpoint: /[siteUrl]/_api/web/lists('[listID]')/GetItems(query=@v1)?@v1={'ViewXml':'<View><Query><Where><Leq><FieldRef Name=Modified/><Value Type=DateTime>[dateStart]</Value></Leq></Where></Query></View>'}&$expand=file
 */
  public async Query4ItemByDateClassic(siteUrl: string, listID: string, defaultUrl: string, dateStart: Date): Promise<MicrosoftGraphBeta.ListItem[]> {
    if (typeof defaultUrl !== "undefined") {
      try {
        // Format the date for SharePoint CAML query (ISO format)
        const formattedDate = dateStart.toISOString();
        // /sites/Demo02/Freigegebene Dokumente/Forms/AllItems.aspx /sites/Demo02/FormServerTemplates/Forms/All Forms.aspx
        /*const temp = defaultUrl.split("/")
        temp.pop();
        //temp.push("ViewForm.aspx?id=");
        temp.push("_layouts/15/listform.aspx?PageType=4&ListId=");
        defaultUrl = temp.join("/");*/
        //https://plumsail.com/docs/forms-sp/how-to/link-to-form.html
        defaultUrl = siteUrl + "/_layouts/15/listform.aspx?PageType=4&ListId=";

        // Construct the ViewXml query
        const viewXml = `<View><Query><Where><Leq><FieldRef Name=Modified/><Value Type=DateTime>${formattedDate}</Value></Leq></Where></Query></View>`;

        const options: ISPHttpClientOptions = {
          headers: {
            'odata-version': '3.0',
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
        //const apiUrl = `${siteUrl}/_api/web/lists('${listID}')/GetItems(query=@v1)?@v1=${queryParam}&$expand=file`;
        const apiUrl = `${siteUrl}/_api/web/lists('${listID}')/GetItems?$expand=ParentList,File,ContentType`;

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
        const items: MicrosoftGraphBeta.ListItem[] = data.d?.results || [];
        items.forEach((item: MicrosoftGraphBeta.ListItem) => {
          // https://[Your SharePoint SiteURL]/_layouts/15/listform.aspx?PageType=[Type]&ListId=[ListGUID]&ID=[Item ID]
          //console.log(item);          
          item.webUrl = `${defaultUrl}${(item as any).ParentList.Id}&id=${(item as any).Id}`;
          //item.webUrl = `/_layouts/15/listform.aspx?PageType=4&ListId=${(item as any).GUID}`
        });
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


