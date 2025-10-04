import * as React from 'react';
import styles from './ContentHealthManager.module.scss';
import type { IContentHealthManagerProps } from './IContentHealthManagerProps';
import { ListView, type IViewField } from '@pnp/spfx-controls-react/lib/ListView';
import { DatePicker, SelectionMode } from '@fluentui/react';
import { SitePicker } from "@pnp/spfx-controls-react/lib/SitePicker";
import type { Site } from '../../../models/Site';
import { Button, Dropdown, Option, Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions, Field } from '@fluentui/react-components';
import GraphDataManager from '../../../services/GraphDataManager';
import { PageProcessing } from '../../../Core/PageProcessing';
import { Page } from '../../../models/Page';
import { PageResult } from '../../../models/PageResult';
import type { LinkInfo } from '../../../models/LinkInfo';
import { SendRegular } from "@fluentui/react-icons";
import { ListInformation } from '../../../models/REST/ListInformation';
import { FieldDateRenderer,FieldTextRenderer } from '@pnp/spfx-controls-react';
import { ListTemplateType } from '../../../Core/ListTemplateTypes';
//import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"

interface IContentHealthManagerState {  
  viewFields: IViewField[];
  //libraryEntries: MicrosoftGraphBeta.List[];
  libraryEntries: ListInformation[];
  pageEntries: Page[];
  SelectedSites: Site[];
  selectedSiteId: string | null;
  pageResults: PageResult[];  
  isReportOpen?: boolean;
  selectedPage?: Page | null;
  dateStartDate: Date | null | undefined;
  isLibraryReportOpen?: boolean;
  selectedLibrary?: ListInformation | null;
}

export default class ContentHealthManager extends React.Component<IContentHealthManagerProps, IContentHealthManagerState> {
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
    { name: 'ContentTypeId', displayName: 'Content Type', sorting: true, isResizable: true, minWidth: 150 }
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
  constructor(props: IContentHealthManagerProps) {
    super(props);

    this.state = {     
      dateStartDate: new Date(),   
      pageResults: [],
      SelectedSites: [],   
      libraryEntries: [],
      selectedSiteId: null,
      isReportOpen: false,
      selectedPage: null,
      isLibraryReportOpen: false,
      selectedLibrary: null,
      viewFields: [
        { name: 'title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 120 },
        { name: 'name', displayName: 'Name', sorting: true, isResizable: true, minWidth: 100 },
        { name: 'webUrl', displayName: 'URL', sorting: false, isResizable: true, minWidth: 200 },
        { name: 'InProgress', displayName: 'InProgress', sorting: false, isResizable: false, minWidth: 50,
          render: (item, index, column) => {  
            return <>
             <SendRegular />
            </>;
            return item.InProgress ? "YES":"NO";
          }
        },        
        { name: 'Links', displayName: 'Links', sorting: false, isResizable: true, minWidth: 200,
          render: (item, index, column) => {                                    
            const entry = this.state.pageResults.filter(x=>x.pageID === item.id)[0];            
            if (typeof entry !== "undefined")
            {
              return `Found ${entry.Links.length}. Broken links: ${entry.Links.filter(x=>x.IsBroken).length}`;
            }
            return "-";
          }
         }
      ],
      pageEntries: []
    };
  }

  private GetLibraryEntryByIndex(index: string):ListInformation
  {    
    return this.state.libraryEntries.filter(x=>x.Id === index)[0];
  }

  public render(): React.ReactElement<IContentHealthManagerProps> {
    return (
      <section className={styles.contentHealthManager}>
        
        <SitePicker
          context={this.props.wpContext}
          label={'Select sites'}
          mode={'site'}
          allowSearch={true}
          multiSelect={true}
          onChange={(sites) => {                
            this.setState({ SelectedSites: sites as Site[] });            
          }}
          placeholder={'Select sites'}
          searchPlaceholder={'Filter sites'} />

        <div className={'ms-Grid'}>
          <div className={'ms-Grid-row'}>
            <div className={'ms-Grid-col ms-sm12 ms-md4 ms-lg3'}>
              <p>TODO</p>
              <label htmlFor={'ddCurrentSite'}>Site selection</label>
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
              <Button onClick={() => this.StartBrokenLinkProcess()}>Find Broken Links</Button>
              <Button onClick={() => this.ShowPageReport()}>Open details</Button>
            </div>
            <div className={'ms-Grid-col ms-sm12 ms-md8 ms-lg9'}>
              <h3>Page library</h3>
              <ListView                
                items={this.state.pageEntries}
                viewFields={this.state.viewFields}
                compact={true}                
                selectionMode={SelectionMode.single}
                selection={this.onListSelectionChanged}/>
              <h3>Site libraries & lists</h3>
              <Field label="Select a date">
                <DatePicker 
                  value={new Date()}
                  minDate={new Date(2000,0,1)}
                  maxDate={new Date()}
                  placeholder="Select a query date..." 
                  onSelectDate={(selectedDate:Date|null) => this.setState(
                    {dateStartDate: selectedDate}
                  )}
                />
              </Field>
              <Button onClick={() => this.StartQueryLstAndLibraries()}>Find old data</Button>
              <Button onClick={() => this.ShowLibraryReport()}>Show details</Button>
              <ListView                
                items={this.state.libraryEntries}
                viewFields={this.viewFieldsLibs}
                compact={true}                
                selectionMode={SelectionMode.single}
                selection={this.onLibrarySelectionChanged}
              />
            </div>            
          </div>
        </div>                
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
                          <div style={{ marginTop: 8, maxHeight: '300px' }}>
                            <ListView                
                              items={this.state.selectedLibrary.FoundItems}
                              viewFields={this.viewFieldsFoundItems}
                              compact={true}                
                              selectionMode={SelectionMode.none}
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

    const dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    const pageAnalyzer = new PageProcessing();
    try {
      // Iterate over all page entries and get their full content
      for (const pageEntry of this.state.pageEntries) {
        try {
          console.log(`Processing page: ${pageEntry.title || pageEntry.name} (ID: ${pageEntry.InProgress})`);
          
          // Get the full page content using GetPageContent method
          const fullPageContent = await dataManager.GetPageContent(this.state.selectedSiteId, pageEntry.id);
                    
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
    const dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    const site : Site = this.GetSelectedSite();
    for (const listInfo of this.state.libraryEntries) {
      const items = await dataManager.Query4ItemByDate(
        site.url,
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

  private async StartQueryLstAndLibraries(): Promise<void> {
    this.CollectItemsFromListAndLibraries();
  }

  private onDropdDownSelectionChanged = async (event: any, data: any): Promise<void> => {    
    const dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    const pages = await dataManager.GetPages4Site(data.optionValue);
    this.setState({ 
      pageEntries: pages,
      selectedSiteId: data.optionValue
    });
    /*const libraries = await dataManager.GetLibraries(data.optionValue);
    console.log(libraries);
    this.setState({ 
      libraryEntries: libraries      
    });*/
    const siteInfo : Site = this.state.SelectedSites.filter(x=>x.id === data.optionValue)[0];    
    const libraries = await dataManager.GetAllLists(siteInfo.url);    
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
    this.setState({ selectedLibrary: this.GetLibraryEntryByIndex(selected!.Id) });
  }

  private GetSelectedSite() : Site
  {
    return this.state.SelectedSites.filter(x=>x.id === this.state.selectedSiteId)[0] as Site;
  }
}
