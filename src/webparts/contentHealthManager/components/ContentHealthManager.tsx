import * as React from 'react';
import styles from './ContentHealthManager.module.scss';
import type { IContentHealthManagerProps } from './IContentHealthManagerProps';
import { ListView, type IViewField } from '@pnp/spfx-controls-react/lib/ListView';
import { SelectionMode } from '@fluentui/react';
import { SitePicker } from "@pnp/spfx-controls-react/lib/SitePicker";
import type { Site } from '../../../models/Site';
import { Button, Dropdown, Option, Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions } from '@fluentui/react-components';
import GraphDataManager from '../../../services/GraphDataManager';
import { PageProcessing } from '../../../Core/PageProcessing';
import { Page } from '../../../models/Page';
import { PageResult } from '../../../models/PageResult';
import type { LinkInfo } from '../../../models/LinkInfo';
import { SendRegular } from "@fluentui/react-icons";
import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"

interface IContentHealthManagerState {  
  viewFields: IViewField[];
  libraryEntries: MicrosoftGraphBeta.List[];
  pageEntries: Page[];
  SelectedSites: Site[];
  selectedSiteId: string | null;
  pageResults: PageResult[];  
  isReportOpen?: boolean;
  selectedPage?: Page | null;
}

export default class ContentHealthManager extends React.Component<IContentHealthManagerProps, IContentHealthManagerState> {
  viewFieldsLibs: IViewField[] = [
    { name: 'displayName', displayName: 'Title', sorting: true, isResizable: true, minWidth: 120, linkPropertyName:'webUrl' },
    { name: 'name', displayName: 'Name', sorting: true, isResizable: true, minWidth: 100 },
    { name: 'lastModifiedDateTime', displayName: 'last change', sorting: true, isResizable: true, minWidth: 100 },
    { name: 'createdDateTime', displayName: 'created at', sorting: true, isResizable: true, minWidth: 100 }
  ];
  constructor(props: IContentHealthManagerProps) {
    super(props);

    this.state = {        
      pageResults: [],
      SelectedSites: [],   
      libraryEntries: [],
      selectedSiteId: null,
      isReportOpen: false,
      selectedPage: null,
      viewFields: [
        { name: 'title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 120, linkPropertyName:'webUrl' },
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
            console.log(entry);
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
            console.log(this.state.SelectedSites);         
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
              <h3>Site libraries</h3>
              <ListView                
                items={this.state.libraryEntries}
                viewFields={this.viewFieldsLibs}
                compact={true}                
                selectionMode={SelectionMode.single}
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
      </section>
    );
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

    const dataManager = new GraphDataManager(this.props.msGraphClientFactory);
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


  public async componentDidMount(): Promise<void> {
    
  }
  private onDropdDownSelectionChanged = async (event: any, data: any): Promise<void> => {    
    const dataManager = new GraphDataManager(this.props.msGraphClientFactory);
    const pages = await dataManager.GetPages4Site(data.optionValue);
    this.setState({ 
      pageEntries: pages,
      selectedSiteId: data.optionValue
    });
    const libraries = await dataManager.GetLibraries(data.optionValue);
    console.log(libraries);
    this.setState({ 
      libraryEntries: libraries      
    });
  }

  private onListSelectionChanged = (items: any[]): void => {
    const selected = (items && items.length > 0) ? (items[0] as Page) : null;
    this.setState({ selectedPage: selected });
  }
}
