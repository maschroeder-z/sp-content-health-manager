import * as React from 'react';
import styles from './ContentHealthManager.module.scss';
import type { IContentHealthManagerProps } from './IContentHealthManagerProps';
import { ListView, type IViewField } from '@pnp/spfx-controls-react/lib/ListView';
import { Checkbox, DatePicker, SelectionMode } from '@fluentui/react';
import { SitePicker } from "@pnp/spfx-controls-react/lib/SitePicker";
import type { Site } from '../../../models/Site';
import { Button, Dropdown, Option, Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions, Field, TabList, Tab, TabValue, Spinner } from '@fluentui/react-components';
import GraphDataManager from '../../../services/GraphDataManager';
import { PageProcessing } from '../../../Core/PageProcessing';
import { Page } from '../../../models/Page';
import { PageResult } from '../../../models/PageResult';
import type { LinkInfo } from '../../../models/LinkInfo';
import { CheckmarkCircleColor, CheckmarkCircleHintRegular, FlagPrideIntersexInclusiveProgressFilled, QuestionCircleColor, WarningColor, Search24Regular, DataTrending24Regular, List24Regular, Link24Regular, Clock24Regular, LockClosed24Regular } from "@fluentui/react-icons";
import { ListInformation } from '../../../models/REST/ListInformation';
import { FieldDateRenderer,FieldTextRenderer } from '@pnp/spfx-controls-react';
import { ListTemplateType } from '../../../Core/ListTemplateTypes';
//import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"

interface IContentHealthManagerState {      
  libraryEntries: ListInformation[];
  pageEntries: Page[];
  SelectedSites: Site[];
  selectedSiteId: string | null;
  pageResults: PageResult[];  
  isReportOpen?: boolean;
  selectedPage?: Page | null;
  dateStartDate: Date |  undefined;
  isLibraryReportOpen?: boolean;
  selectedLibrary?: ListInformation | null;
  selectedTabValue: TabValue;
  chkShowLists: boolean;
  chkShowLibaries: boolean;
  selectedFoundItem?: any | null;
  isQueryingLibraries?: boolean;
}

export default class ContentHealthManager extends React.Component<IContentHealthManagerProps, IContentHealthManagerState> {
  tempSelectedSites : Site[] =   [
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
]
  dataManager: GraphDataManager;
  // View fields for found items in library report dialog
  viewFieldsFoundItems: IViewField[] = [
    { name: 'Id', displayName: 'ID', sorting: true, isResizable: true, minWidth: 80, linkPropertyName:'webUrl' },
    { name: 'Title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 200 },
    { 
      name: 'Created', displayName: 'Created', sorting: true, isResizable: true, minWidth: 120,
      render: (item: any, index, column) => {
        const date = new Date(item.Created);
        return <FieldDateRenderer text={date.toLocaleDateString()} />;    
      }
    },
    { 
      name: 'Modified', displayName: 'Modified', sorting: true, isResizable: true, minWidth: 120,
      render: (item: any, index, column) => {
        const date = new Date(item.Modified);
        return <FieldDateRenderer text={date.toLocaleDateString()} />;    
      }
    },
    { name: 'ContentTypeId', displayName: 'Content Type', sorting: true, isResizable: true, minWidth: 150,
      render: (item: any, inxdex, column) => {
        if (typeof item.ContentType !== "undefined")
          return item.ContentType;              
        return item["ContentType.Name"];
      }
     }
  ];

  // BaseTemplate BaseType EnableAttachments EnableFolderCreation EnableVersioning ForceCheckout ItemCount LastItemModifiedDate LastItemUserModifiedDate
  viewFieldsLibs: IViewField[] = [
    { name: 'Title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 120, linkPropertyName:'DefaultView.ServerRelativeUrl'},
    { 
      name: 'BaseTemplate', displayName: 'Template', sorting: true, isResizable: true, minWidth: 100,
      render: (item:ListInformation, index, column) => {        
        return ListTemplateType[item.BaseTemplate];
      }
    },
    { 
      name: 'Created', displayName: 'created at', sorting: true, isResizable: true, minWidth: 100,
      render: (item:ListInformation, index, column) => {
        const date = new Date(item.Created);
        return <FieldDateRenderer text={date.toLocaleDateString()} />;    
      }
    },
    { 
      name: 'LastItemModifiedDate', displayName: 'Last change', sorting: true, isResizable: true, minWidth: 120, linkPropertyName:'webUrl',
      render: (item:ListInformation, index, column) => {
        const date = new Date(item.LastItemModifiedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;  
      }
    },
    { 
      name: 'LastItemUserModifiedDate', displayName: 'User changed', sorting: true, isResizable: true, minWidth: 120, linkPropertyName:'webUrl',
      render: (item:ListInformation, index, column) => {
        const date = new Date(item.LastItemUserModifiedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;
      }
    },
    { 
      name: 'LastItemDeletedDate', displayName: 'last deletion', sorting: true, isResizable: true, minWidth: 100,
      render: (item:ListInformation, index, column) => {
        const date = new Date(item.LastItemDeletedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;
      }
    },
    { name: 'ItemCount', displayName: 'Items', sorting: true, isResizable: true, minWidth: 120 },
    { name: 'FoundItems', displayName: 'Found', sorting: true, isResizable: true, minWidth: 120,
      render: (item:ListInformation, index, column) => {             
        const entry = this.GetLibraryEntryByIndex(item.Id);
        if (typeof entry.FoundItems !== "undefined" && entry.FoundItems !== null)
        {
          return <FieldTextRenderer text={`Found: ${entry.FoundItems?.length}`} />;
        }
        else
          return <FieldTextRenderer text="start query fo results" />;
      }
     },
    { name: 'Description', displayName: 'Description', sorting: true, isResizable: true, minWidth: 100 }
  ];

  viewFieldsPage: IViewField[] = [
    { name: 'title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 120 },
    { name: 'name', displayName: 'Name', sorting: true, isResizable: true, minWidth: 100 },
    { name: 'webUrl', displayName: 'URL', sorting: false, isResizable: true, minWidth: 200 },     
    { name: 'Links', displayName: 'Links', sorting: false, isResizable: true, minWidth: 200,
      render: (item, index, column) => {                                    
        const entry = this.state.pageResults.filter(x=>x.pageID === item.id)[0];            

        if (typeof entry === "undefined" || typeof entry.Links === "undefined")
        {
          return <>          
          <CheckmarkCircleHintRegular />
          </>;
        }

        if (entry.Links.filter(x=>x.IsBroken).length>0)
        {
          return (<>
            <WarningColor />
            &nbsp;<span>Found {entry.Links.length}. Broken links: {entry.Links.filter(x=>x.IsBroken).length}</span>
            </>);
        }
        return <>          
          <CheckmarkCircleColor />
          &nbsp;
          <span>Found {entry.Links.length}. Broken links: {entry.Links.filter(x=>x.IsBroken).length}</span>
          </>; 
      }
     }
  ];

  constructor(props: IContentHealthManagerProps) {
    super(props);

    this.state = {     
      dateStartDate: new Date(),   
      pageResults: [],
      SelectedSites: this.tempSelectedSites,   
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
      isQueryingLibraries: false
    };
    this.dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
  }

  private GetLibraryEntryByIndex(index: string):ListInformation
  {    
    return this.state.libraryEntries.filter(x=>x.Id === index)[0];
  }
/**https://storybooks.fluentui.dev/react/?path=/docs/components-tablist--docs*/
  public render(): React.ReactElement<IContentHealthManagerProps> {
    return (
      <section className={styles.contentHealthManager}>
        {this.state.SelectedSites.length === 0 && (
          <div className={styles.summarySection}>
            <div className={styles.summaryDescription}>
              <Search24Regular className={styles.summaryIcon} />
              <div>
                <h3>Content Health Manager</h3>
                <p>
                  <DataTrending24Regular className={styles.inlineIcon} />
                  This application helps you analyze and maintain the health of your SharePoint content by identifying broken links, 
                  finding outdated content based on modification dates, and detecting checked-out items that may need attention. 
                  Use this tool to ensure your SharePoint sites remain organized and accessible.
                </p>
              </div>
            </div>
            
            <div className={styles.instructionsSection}>
              <h4><List24Regular className={styles.inlineIcon} />How to use:</h4>
              <ol className={styles.stepList}>
                <li>
                  <strong>First, select sites</strong> - Choose all the SharePoint sites you want to analyze from the site picker below.
                </li>
                <li>
                  <strong>Second, select a single site</strong> - From the dropdown, pick one specific site to process and analyze.
                </li>
                <li>
                  <strong>Start query to find:</strong>
                  <ul className={styles.subList}>
                    <li><Link24Regular className={styles.inlineIcon} />Broken links in pages</li>
                    <li><Clock24Regular className={styles.inlineIcon} />Old content for a given start date</li>
                    <li><LockClosed24Regular className={styles.inlineIcon} />Checked out content items</li>
                  </ul>
                </li>
              </ol>
            </div>
          </div>
        )}

        <div className={styles.row}>
          <div className={styles['col-sm12']}>
            {this.state.SelectedSites.length === 0 && <p className={styles.infoMessage}><QuestionCircleColor />Select first all sites that you want to query and process</p>}            
            <Field label="Select sites">
              <SitePicker
                context={this.props.wpContext}              
                mode={'site'}         
                selectedSites={this.tempSelectedSites}       
                allowSearch={true}
                multiSelect={true}
                className={styles.sitePicker}                
                trimDuplicates={true}
                onChange={(sites) => {                
                  console.log(sites);
                  this.setState({ SelectedSites: sites as Site[] });            
                }}
                placeholder={'Select all sites to process...'}
                searchPlaceholder={'Filter sites'} />
              </Field>
          </div>
          <div className={styles['col-sm12']}>      
            {this.state.SelectedSites.length > 0 && this.state.selectedSiteId === null && <div>
              <p className={styles.infoMessage}><QuestionCircleColor />To continue, please select first a site to process</p>
            </div>}
            {this.state.SelectedSites.length > 0 &&
              <Field label="Choose a Site">
                <Dropdown 
                  id={'ddCurrentSite'} 
                  inlinePopup={true}                 
                  onOptionSelect={this.onDropdDownSelectionChanged}
                  placeholder="Select a Site to process">
                  {this.state.SelectedSites.map((entry:Site) => (
                    <Option value={entry.id} key={entry.webId} >
                      {entry.title}
                    </Option>
                  ))}
                </Dropdown>     
              </Field>
              }
          </div>
        </div>

        {this.state.selectedSiteId && <>        
        <p className={styles.infoMessage}><FlagPrideIntersexInclusiveProgressFilled />Results for site: 
          <a href={this.GetSelectedSite().url} target={'_blank'} rel={'noreferrer'}><strong>{this.GetSelectedSite().title}</strong></a>
        </p>
        <TabList selectedValue={this.state.selectedTabValue} onTabSelect={this.onTabSelect}>
          <Tab value="tab1">Broken Links Analysis</Tab>
          <Tab value="tab2">Library Analysis</Tab>
        </TabList> </> }

        {this.state.selectedTabValue === 'tab2' && (
          <div id="Register1" className={styles.row}>            
            <div className={`${styles.row} ${styles.libraryCommands}`}> 
              <div className={styles['col-sm5']}>    
                <Field label="Select a date" orientation="horizontal" >
                  <DatePicker 
                    value={this.state.dateStartDate}
                    minDate={new Date(2000,0,1)}
                    maxDate={new Date()}
                    placeholder="Select a query date..." 
                    onSelectDate={(selectedDate:Date|undefined) => this.setState(
                      {dateStartDate: selectedDate}
                    )}
                  />
                </Field>
              </div>
              <div className={`${styles['col-sm7']} ${styles.libraryCommandsLeft}`}>    
                <Button onClick={() => this.StartQueryLstAndLibraries()} disabled={this.state.isQueryingLibraries}>Search</Button>
                {this.state.isQueryingLibraries && <Spinner size="tiny" className={styles.progressSpinner} />}
                &nbsp;
                <Button onClick={() => this.StartQueryCheckedOutItems()}>Checked-out items</Button>                                
              </div>
            </div>
            <div className={`${styles.row} ${styles.libraryCommands}`}> 
                <div className={styles['col-sm4']}>
                  <Button onClick={() => this.ShowLibraryReport()} disabled={!this.state.selectedLibrary}>Open details</Button>
                </div>
                
                <div className={`${styles['col-sm8']} ${styles.checkboxContainer}`}>
                  <Checkbox
                    checked={this.state.chkShowLibaries}
                    onChange={async (ev, checked: boolean) => {                                              
                        const libraries = await this.dataManager.GetAllLists(this.GetSelectedSite().url, this.state.chkShowLists, checked);
                        this.setState({ 
                          libraryEntries: libraries,
                          chkShowLibaries: checked
                        }); 
                      }
                    }
                    label="Libraries"
                  />
                  <Checkbox 
                    checked={this.state.chkShowLists}
                    onChange={async (ev, checked: boolean) => {                    
                        const libraries = await this.dataManager.GetAllLists(this.GetSelectedSite().url, checked, this.state.chkShowLibaries);
                        this.setState({ 
                          libraryEntries: libraries,
                          chkShowLists: checked
                        });                        
                      }
                    }
                    label="Lists"
                  />  
                </div>
            </div>
            <ListView                
              items={this.state.libraryEntries}
              viewFields={this.viewFieldsLibs}
              compact={true}                
              selectionMode={SelectionMode.single}
              selection={this.onLibrarySelectionChanged} />
          </div>
        )}

        {this.state.selectedTabValue === 'tab1' && (
          <div id="Register2" className={styles.row}>
          <div className={`${styles.row} ${styles.libraryCommands}`}> 
            <div className={`${styles['col-sm12']} ${styles.libraryCommandsLeft}`}>              
              <Button onClick={() => this.StartBrokenLinkProcess()}>Find Broken Links</Button>
              &nbsp;
              <Button onClick={() => this.ShowPageReport()} disabled={!this.state.selectedPage}>Open details</Button>
            </div>
          </div>
          <ListView                
            items={this.state.pageEntries}
            viewFields={this.viewFieldsPage}
            compact={true}                
            selectionMode={SelectionMode.single}
            selection={this.onListSelectionChanged}/>              
          </div>
        )}                
                
        <Dialog open={!!this.state.isReportOpen} onOpenChange={(_: any, data: any) => this.setState({ isReportOpen: !!data.open })} modalType={'alert'}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Page report</DialogTitle>
              <DialogContent style={{ padding: 12 }}>
                {this.state.selectedPage ? (
                  <div>
                    <div><strong>Title:</strong> {this.state.selectedPage.title || this.state.selectedPage.name}</div>
                    <div><strong>URL:</strong> <a href={this.state.selectedPage.webUrl} target={'_blank'} rel={'noreferrer'}>{this.state.selectedPage.webUrl}</a></div>
                    {(() => {
                      const entry = this.state.pageResults.filter((x: PageResult) => x.pageID === this.state.selectedPage!.id)[0];
                      if (entry) {
                        return (
                          <div style={{ marginTop: 8 }}>
                            <div><strong>Total links:</strong> {entry.Links.length}</div>
                            <div><strong>Broken links:</strong> {entry.Links.filter((l: LinkInfo) => l.IsBroken).length}</div>
                            <div style={{ marginTop: 12 }}>
                              <div><strong>All Links:</strong></div>
                              <div style={{ maxHeight: '300px', overflowY: 'auto', marginTop: 8, border: '1px solid #ccc', padding: 8 }}>
                                {entry.Links.length > 0 ? (
                                  entry.Links.map((link: LinkInfo, index: number) => (
                                    <div key={index} style={{ 
                                      padding: '8px', 
                                      marginBottom: '4px', 
                                      border: '1px solid #e0e0e0', 
                                      borderRadius: '4px',
                                      backgroundColor: link.IsBroken ? '#ffebee' : '#f5f5f5'
                                    }}>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                        <span style={{ 
                                          color: link.IsBroken ? '#d32f2f' : '#2e7d32', 
                                          fontWeight: 'bold',
                                          fontSize: '12px'
                                        }}>
                                          {link.IsBroken ? '❌ BROKEN' : '✅ OK'}
                                        </span>
                                      </div>
                                      <div style={{ marginTop: '4px' }}>
                                        <div><strong>Title:</strong> {link.title || 'No title'}</div>
                                        <div><strong>URL:</strong> 
                                          <a href={link.url} target="_blank" rel="noopener noreferrer" style={{ marginLeft: '4px', color: '#0078d4' }}>
                                            {link.url}
                                          </a>
                                        </div>
                                      </div>
                                    </div>
                                  ))
                                ) : (
                                  <div style={{ padding: '8px', color: '#666', fontStyle: 'italic' }}>
                                    No links found on this page.
                                  </div>
                                )}
                              </div>
                            </div>
                          </div>
                        );
                      }
                      return <div style={{ marginTop: 8 }}>No link analysis available.</div>;
                    })()}
                  </div>
                ) : (
                  <div>No item selected.</div>
                )}
              </DialogContent>
              <DialogActions>
                <Button appearance={'secondary'} onClick={() => this.setState({ isReportOpen: false })}>Close</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>

        <Dialog open={!!this.state.isLibraryReportOpen} onOpenChange={(_: any, data: any) => this.setState({ isLibraryReportOpen: !!data.open })} modalType={'alert'}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Library report</DialogTitle>
              <DialogContent style={{ padding: 12 }}>
                {this.state.selectedLibrary ? (
                  <div>
                    <div><strong>Title:</strong> {this.state.selectedLibrary.Title || 'N/A'}</div>
                    <div><strong>Template:</strong> {ListTemplateType[this.state.selectedLibrary.BaseTemplate] || 'N/A'}</div>
                    <div><strong>Description:</strong> {this.state.selectedLibrary.Description || 'N/A'}</div>
                    <div><strong>Item Count:</strong> {this.state.selectedLibrary.ItemCount}</div>
                    <div><strong>Created:</strong> {new Date(this.state.selectedLibrary.Created).toLocaleDateString()}</div>
                    <div><strong>Last Modified:</strong> {new Date(this.state.selectedLibrary.LastItemModifiedDate).toLocaleString()}</div>
                    <div><strong>Last User Modified:</strong> {new Date(this.state.selectedLibrary.LastItemUserModifiedDate).toLocaleString()}</div>
                    {this.state.selectedLibrary.LastItemDeletedDate && (
                      <div><strong>Last Deleted:</strong> {new Date(this.state.selectedLibrary.LastItemDeletedDate).toLocaleString()}</div>
                    )}
                    <div><strong>Enable Versioning:</strong> {this.state.selectedLibrary.EnableVersioning ? 'Yes' : 'No'}</div>
                    <div><strong>Enable Attachments:</strong> {this.state.selectedLibrary.EnableAttachments ? 'Yes' : 'No'}</div>
                    <div><strong>Enable Folder Creation:</strong> {this.state.selectedLibrary.EnableFolderCreation ? 'Yes' : 'No'}</div>
                    
                    <div style={{ marginTop: 16 }}>
                      <h4>Overview list entries</h4>
                      {this.state.selectedLibrary.FoundItems && this.state.selectedLibrary.FoundItems.length > 0 ? (
                        <div>
                          <div><strong>Total items found:</strong> {this.state.selectedLibrary.FoundItems.length}</div>
                          <Button 
                            onClick={this.onShowPermissionsClick}
                            disabled={!this.state.selectedFoundItem}
                            appearance="secondary"
                            style={{ marginBottom: '8px' }}
                          >
                            Show Permissions
                          </Button>
                          <div style={{ marginTop: 8, maxHeight: '300px' }}>
                            <ListView                
                              items={this.state.selectedLibrary.FoundItems}
                              viewFields={this.viewFieldsFoundItems}
                              compact={true}                
                              selectionMode={SelectionMode.single}
                              selection={this.onFoundItemSelectionChanged}
                            />
                          </div>
                        </div>
                      ) : (
                        <div style={{ padding: '16px', backgroundColor: '#f5f5f5', border: '1px solid #ddd', borderRadius: '4px', textAlign: 'center' }}>
                          <p style={{ margin: 0, color: '#666' }}>Query the library for results</p>
                        </div>
                      )}
                    </div>
                  </div>
                ) : (
                  <div>No library selected.</div>
                )}
              </DialogContent>
              <DialogActions>
                <Button appearance={'secondary'} onClick={() => this.setState({ isLibraryReportOpen: false })}>Close</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </section>
    );
  }

  public async componentDidMount(): Promise<void> {
    
  }

  private ShowLibraryReport():void
  {
    if (!this.state.selectedLibrary) {
      return;
    }
    this.setState({ isLibraryReportOpen: true });
  }

  private ShowPageReport():void
  {
    if (!this.state.selectedPage) {
      return;
    }
    this.setState({ isReportOpen: true });
  }

  private async StartBrokenLinkProcess(): Promise<void>
  {       
    if (!this.state.selectedSiteId) {
      console.warn('No site selected. Please select a site first.');
      return;
    }

    if (!this.state.pageEntries || this.state.pageEntries.length === 0) {
      console.warn('No pages found for the selected site.');
      return;
    }

    console.log(`Starting broken link process for site: ${this.state.selectedSiteId}`);
    console.log(`Processing ${this.state.pageEntries.length} pages...`);

    //const dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    const pageAnalyzer = new PageProcessing();
    try {
      // Iterate over all page entries and get their full content
      for (const pageEntry of this.state.pageEntries) {
        try {
          console.log(`Processing page: ${pageEntry.title || pageEntry.name} (ID: ${pageEntry.InProgress})`);
          
          // Get the full page content using GetPageContent method
          const fullPageContent = await this.dataManager.GetPageContent(this.state.selectedSiteId, pageEntry.id);
                    
          // TODO: Add broken link detection logic here
          const resultLinks = await pageAnalyzer.AnalyzePageContent(fullPageContent.canvasLayout!);          
          this.state.pageResults.push({
            pageID: pageEntry.id,
            Links: resultLinks!
          });
          
          this.setState({
            pageEntries: this.state.pageEntries
          })
          
        } catch (error) {
          console.error(`Error processing page ${pageEntry.title || pageEntry.name}:`, error);
        }
      }            
      /*this.setState({
        pageEntries: this.state.pageEntries
      })*/
    } catch (error) {
      console.error('Error during broken link process:', error);
    }
  }

  public async CollectItemsFromListAndLibraries():Promise<void>
  {
    const site : Site = this.GetSelectedSite();
    for (const listInfo of this.state.libraryEntries) {
      const items = await this.dataManager.Query4ItemByDate(
        site,
        listInfo.Id,        
        listInfo.DefaultView.ServerRelativeUrl,
        this.state.dateStartDate!
      );
      listInfo.FoundItems = items;            
      this.setState({ 
        libraryEntries: this.state.libraryEntries      
      });   
    }
  }

  public async GetCheckedOutItems():Promise<void>
  {
    const site : Site = this.GetSelectedSite();
    for (const listInfo of this.state.libraryEntries) {
      const items = await this.dataManager.Query4CheckedOutItems(
        site,
        listInfo.Id,        
        listInfo.DefaultView.ServerRelativeUrl,
        this.state.dateStartDate!
      );
      listInfo.FoundItems = items;                  
      this.setState({ 
        libraryEntries: this.state.libraryEntries      
      });   
    }
  }

  private async StartQueryCheckedOutItems(): Promise<void> {
    this.GetCheckedOutItems();
  }

  public async GetPermission4SelectedItem(site: Site, listID: string, listItemID: string): Promise<void> {
    try {
      const permissions = await this.dataManager.GetPermission4Item(site, listID, listItemID);
      console.log('Item permissions:', permissions);
      // You can add additional logic here to handle the permissions data
      // For example, display them in a dialog or update the UI state
    } catch (error) {
      console.error('Error retrieving item permissions:', error);
    }
  }

  private async StartQueryLstAndLibraries(): Promise<void> {
    this.setState({ isQueryingLibraries: true });
    try {
      await this.CollectItemsFromListAndLibraries();
    } finally {
      this.setState({ isQueryingLibraries: false });
    }
  }

  private onDropdDownSelectionChanged = async (event: any, data: any): Promise<void> => {    
    const dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    const pages = await dataManager.GetPages4Site(data.optionValue);
    this.setState({ 
      selectedTabValue: this.state.selectedTabValue === null ? "tab1":this.state.selectedTabValue,
      pageEntries: pages,
      selectedSiteId: data.optionValue
    });
    const siteInfo : Site = this.state.SelectedSites.filter(x=>x.id === data.optionValue)[0];    
    const libraries = await dataManager.GetAllLists(siteInfo.url, this.state.chkShowLists, this.state.chkShowLibaries);
    console.log("All lists", libraries);
    this.setState({ 
      libraryEntries: libraries      
    });    
  }

  private onListSelectionChanged = (items: any[]): void => {
    const selected = (items && items.length > 0) ? (items[0] as Page) : null;
    this.setState({ selectedPage: selected });
  }

  private onLibrarySelectionChanged = (items: any[]): void => {
    const selected = (items && items.length > 0) ? (items[0] as ListInformation) : null;
    if (selected !== null)
      this.setState({ selectedLibrary: this.GetLibraryEntryByIndex(selected!.Id) });
    else
    this.setState({ selectedLibrary: null });
  }

  private GetSelectedSite() : Site
  {
    return this.state.SelectedSites.filter(x=>x.id === this.state.selectedSiteId)[0] as Site;
  }

  private onTabSelect = (event: any, data: { value: TabValue }): void => {
    this.setState({ selectedTabValue: data.value });
  }

  private onFoundItemSelectionChanged = (items: any[]): void => {
    const selected = (items && items.length > 0) ? items[0] : null;
    this.setState({ selectedFoundItem: selected });
  }

  private onShowPermissionsClick = async (): Promise<void> => {
    if (!this.state.selectedFoundItem || !this.state.selectedLibrary) {
      console.warn('No item selected or no library selected');
      return;
    }

    const site = this.GetSelectedSite();    
    await this.GetPermission4SelectedItem(site, this.state.selectedLibrary.Id, this.state.selectedFoundItem.Id);
  }
}
