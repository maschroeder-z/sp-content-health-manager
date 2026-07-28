import * as React from 'react';
import styles from './ContentHealthManager.module.scss';
import type { IContentHealthManagerProps } from './IContentHealthManagerProps';
import { ListView, type IViewField } from '@pnp/spfx-controls-react/lib/ListView';
import { Checkbox, DatePicker, SelectionMode, Toggle } from '@fluentui/react';
import { SitePicker } from "@pnp/spfx-controls-react/lib/SitePicker";
import type { Site } from '../../../models/Site';
import { Button, Dropdown, Option, Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions, Field, TabList, Tab, TabValue, Spinner, Tooltip } from '@fluentui/react-components';
import GraphDataManager from '../../../services/GraphDataManager';
import { PageProcessing } from '../../../Core/PageProcessing';
import { Page } from '../../../models/Page';
import { PageResult } from '../../../models/PageResult';
import type { LinkInfo } from '../../../models/LinkInfo';
import { CheckmarkCircleColor, CheckmarkCircleHintRegular, FlagPrideIntersexInclusiveProgressFilled, QuestionCircleColor, WarningColor, Search24Regular, DataTrending24Regular, List24Regular, Link24Regular, Clock24Regular, LockClosed24Regular, ChevronDown24Regular, ChevronUp24Regular, DatabaseSearch24Regular, Open24Regular, Dismiss24Regular, KeyMultiple24Regular, Info24Regular, PeopleTeam16Regular, Person16Regular } from "@fluentui/react-icons";
import { ListInformation } from '../../../models/REST/ListInformation';
import PermissionsManager from '../../../services/PermissionsManager';
import { SharePointPrincipalPermission } from '../../../models/REST/Permissions';
import { FieldDateRenderer,FieldTextRenderer } from '@pnp/spfx-controls-react';
import { ListTemplateType } from '../../../Core/ListTemplateTypes';
import * as strings from 'ContentHealthManagerWebPartStrings';
//import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"

interface IContentHealthManagerState {      
  libraryEntries: ListInformation[];
  pageEntries: Page[];
  SelectedSites: Site[];
  selectedSiteId: string | null;
  pageResults: PageResult[];  
  isReportOpen?: boolean;
  selectedPage?: Page | null;
  dateStartDate: Date |  undefined | null;
  isLibraryReportOpen?: boolean;
  selectedLibrary?: ListInformation | null;
  selectedTabValue: TabValue;
  chkShowLists: boolean;
  chkShowLibaries: boolean;
  selectedFoundItem?: any | null;
  isQueryingLibraries?: boolean;
  isFilteringLibraries?: boolean;
  isProcessingBrokenLinks?: boolean;
  expandedContentSections: Set<string>;
  showOnlyBrokenLinks: boolean;
  isPagePermissionsOpen?: boolean;
  pagePermissions: SharePointPrincipalPermission[];
  isLoadingPagePermissions?: boolean;
  pagePermissionsError?: string | null;
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
  permissionsManager: PermissionsManager;
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
    { name: 'ItemCount', displayName: 'Items', sorting: true, isResizable: true, minWidth: 120 },
    { name: 'FoundItems', displayName: strings.FoundLabel, sorting: true, isResizable: true, minWidth: 120,
      render: (item:ListInformation, index, column) => {             
        const entry = this.GetLibraryEntryByIndex(item.Id);
        if (typeof entry.FoundItems !== "undefined" && entry.FoundItems !== null)
        {
          return <FieldTextRenderer text={`${strings.FoundLabel}: ${entry.FoundItems?.length}`} />;
        }
        else
          return <FieldTextRenderer text={strings.StartQueryForResults} />;
      }
     },    
    { 
      name: 'Created', displayName: strings.CreatedAtLabel, sorting: true, isResizable: true, minWidth: 100,
      render: (item:ListInformation, index, column) => {
        const date = new Date(item.Created);
        return <FieldDateRenderer text={date.toLocaleDateString()} />;    
      }
    },
    { 
      name: 'LastItemModifiedDate', displayName: strings.LastChangeLabel, sorting: true, isResizable: true, minWidth: 120, linkPropertyName:'webUrl',
      render: (item:ListInformation, index, column) => {
        const date = new Date(item.LastItemModifiedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;  
      }
    },
    { 
      name: 'LastItemUserModifiedDate', displayName: strings.UserChangedLabel, sorting: true, isResizable: true, minWidth: 120, linkPropertyName:'webUrl',
      render: (item:ListInformation, index, column) => {
        const date = new Date(item.LastItemUserModifiedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;
      }
    },
    { 
      name: 'LastItemDeletedDate', displayName: strings.LastDeletionLabel, sorting: true, isResizable: true, minWidth: 100,
      render: (item:ListInformation, index, column) => {
        const date = new Date(item.LastItemDeletedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;
      }
    },
    { name: 'ItemCount', displayName: 'Items', sorting: true, isResizable: true, minWidth: 120 },
    { name: 'FoundItems', displayName: strings.FoundLabel, sorting: true, isResizable: true, minWidth: 120,
      render: (item:ListInformation, index, column) => {             
        const entry = this.GetLibraryEntryByIndex(item.Id);
        if (typeof entry.FoundItems !== "undefined" && entry.FoundItems !== null)
        {
          return <FieldTextRenderer text={`${strings.FoundLabel}: ${entry.FoundItems?.length}`} />;
        }
        else
          return <FieldTextRenderer text={strings.StartQueryForResults} />;
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
            &nbsp;<span>{strings.FoundLinksCount.replace('{0}', entry.Links.length.toString()).replace('{1}', entry.Links.filter(x=>x.IsBroken).length.toString())}</span>
            </>);
        }
        return <>          
          <CheckmarkCircleColor />
          &nbsp;
          <span>{strings.FoundLinksCount.replace('{0}', entry.Links.length.toString()).replace('{1}', entry.Links.filter(x=>x.IsBroken).length.toString())}</span>
          </>; 
      }
     }
  ];

  viewFieldsPermissions: IViewField[] = [
    { name: 'displayName', displayName: strings.PrincipalNameLabel, sorting: true, isResizable: true, minWidth: 180 },
    { name: 'isGroup', displayName: strings.PrincipalTypeLabel, sorting: true, isResizable: true, minWidth: 100,
      render: (item: SharePointPrincipalPermission) => (
        <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
          {item.isGroup ? <PeopleTeam16Regular /> : <Person16Regular />}
          <span>{item.isGroup ? strings.GroupLabel : strings.UserLabel}</span>
        </span>
      )
    },
    { name: 'loginName', displayName: strings.LoginNameLabel, sorting: false, isResizable: true, minWidth: 220 },
    { name: 'roles', displayName: strings.RolesLabel, sorting: false, isResizable: true, minWidth: 200,
      render: (item: SharePointPrincipalPermission) => <FieldTextRenderer text={(item.roles || []).join(', ')} />
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
      isQueryingLibraries: false,
      isFilteringLibraries: false,
      isProcessingBrokenLinks: false,
      expandedContentSections: new Set<string>(),
      showOnlyBrokenLinks: false,
      isPagePermissionsOpen: false,
      pagePermissions: [],
      isLoadingPagePermissions: false,
      pagePermissionsError: null
    };
    this.dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    this.permissionsManager = new PermissionsManager(this.props.spHTTPClient);
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
                <h3>{strings.ContentHealthManagerTitle}</h3>
                <p>
                  <DataTrending24Regular className={styles.inlineIcon} />
                  {strings.ContentHealthManagerDescription}
                </p>
              </div>
            </div>
            
            <div className={styles.instructionsSection}>
              <h4><List24Regular className={styles.inlineIcon} />{strings.HowToUseHeading}</h4>
              <ol className={styles.stepList}>
                <li>
                  <strong>{strings.FirstSelectSites.split(' - ')[0]}</strong> - {strings.FirstSelectSites.split(' - ')[1]}
                </li>
                <li>
                  <strong>{strings.SecondSelectSingleSite.split(' - ')[0]}</strong> - {strings.SecondSelectSingleSite.split(' - ')[1]}
                </li>
                <li>
                  <strong>{strings.StartQueryToFind}</strong>
                  <ul className={styles.subList}>
                    <li><Link24Regular className={styles.inlineIcon} />{strings.BrokenLinksInPages}</li>
                    <li><Clock24Regular className={styles.inlineIcon} />{strings.OldContentForDate}</li>
                    <li><LockClosed24Regular className={styles.inlineIcon} />{strings.CheckedOutContentItems}</li>
                  </ul>
                </li>
              </ol>
            </div>
          </div>
        )}

        <div className={styles.row}>
          <div className={styles['col-sm12']}>
            {this.state.SelectedSites.length === 0 && <p className={styles.infoMessage}><QuestionCircleColor />{strings.SelectFirstAllSites}</p>}            
            <Field label={strings.SelectSitesLabel}>
              <SitePicker
                context={this.props.wpContext as any}              
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
                placeholder={strings.SelectAllSitesPlaceholder}
                searchPlaceholder={strings.FilterSitesPlaceholder} />
              </Field>
          </div>
          <div className={styles['col-sm12']}>      
            {this.state.SelectedSites.length > 0 && this.state.selectedSiteId === null && <div>
              <p className={styles.infoMessage}><QuestionCircleColor />{strings.ToContinueSelectSite}</p>
            </div>}
            {this.state.SelectedSites.length > 0 &&
              <Field label={strings.ChooseSiteLabel}>
                <Dropdown 
                  id={'ddCurrentSite'} 
                  inlinePopup={true}                 
                  onOptionSelect={this.onDropdDownSelectionChanged}
                  placeholder={strings.SelectSitePlaceholder}>
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
        <p className={styles.infoMessage}><FlagPrideIntersexInclusiveProgressFilled />{strings.ResultsForSite} 
          <a href={this.GetSelectedSite().url} target={'_blank'} rel={'noreferrer'}><strong>{this.GetSelectedSite().title}</strong></a>
        </p>
        <TabList selectedValue={this.state.selectedTabValue} onTabSelect={this.onTabSelect}>
          <Tab value="tab1">{strings.BrokenLinksAnalysisTab}</Tab>
          <Tab value="tab2">{strings.LibraryAnalysisTab}</Tab>
        </TabList> </> }

        {this.state.selectedTabValue === 'tab2' && (
          <div id="Register1" className={styles.row}>
            <div className={styles.row}>
              <div className={styles['col-sm12']}>
                <div className={styles.noteBox}>
                  <Info24Regular className={styles.noteBoxIcon} />
                  <span>{strings.SelectDateHint}</span>
                </div>
              </div>
            </div>
            <div className={`${styles.row} ${styles.libraryCommands}`}>
              <div className={styles['col-sm5']}>
                <Field label={strings.SelectDateLabel} orientation="horizontal">
                  <DatePicker
                    value={this.state.dateStartDate as Date | undefined}
                    minDate={new Date(2000,0,1)}
                    maxDate={new Date()}
                    placeholder={strings.SelectQueryDatePlaceholder} 
                    onSelectDate={(selectedDate:Date|undefined|null) => this.setState(
                      {dateStartDate: selectedDate}
                    )}
                  />
                </Field>
              </div>
              <div className={`${styles['col-sm7']} ${styles.libraryCommandsLeft}`}>
                <Tooltip
                  content={this.state.selectedLibrary ? strings.TooltipQueryLibrary : strings.TooltipQueryAllLibraries}
                  relationship="label">
                  <Button icon={<DatabaseSearch24Regular />} onClick={() => this.StartQueryLstAndLibraries()} disabled={this.state.isQueryingLibraries}>
                    {!this.state.selectedLibrary && <span>{strings.QueryAllLibraries}</span>}
                    {this.state.selectedLibrary && <span>{strings.QueryLibrary}</span>}
                  </Button>
                </Tooltip>


                {this.state.isQueryingLibraries && <Spinner size="tiny" className={styles.progressSpinner} />}
                &nbsp;
                <Tooltip content={strings.TooltipCheckedOutItems} relationship="label">
                  <Button icon={<LockClosed24Regular />} onClick={() => this.StartQueryCheckedOutItems()}>{strings.CheckedOutItems}</Button>
                </Tooltip>
              </div>
            </div>
            <div className={`${styles.row} ${styles.libraryCommands}`}>
                <div className={styles['col-sm4']}>
                  <Tooltip content={strings.TooltipOpenLibraryDetails} relationship="label">
                    <Button icon={<Open24Regular />} onClick={() => this.ShowLibraryReport()} disabled={!this.state.selectedLibrary}>{strings.OpenDetails}</Button>
                  </Tooltip>
                </div>
                
                <div className={`${styles['col-sm8']} ${styles.checkboxContainer}`}>
                  <Checkbox
                    checked={this.state.chkShowLibaries}
                    disabled={this.state.isFilteringLibraries}
                    onChange={(ev, checked: boolean | undefined) => {
                        void this.UpdateLibraryFilter(checked || false, this.state.chkShowLists);
                      }
                    }
                    label={strings.LibrariesCheckbox}
                  />
                  <Checkbox
                    checked={this.state.chkShowLists}
                    disabled={this.state.isFilteringLibraries}
                    onChange={(ev, checked: boolean | undefined) => {
                        void this.UpdateLibraryFilter(this.state.chkShowLibaries, checked || false);
                      }
                    }
                    label={strings.ListsCheckbox}
                  />
                  {this.state.isFilteringLibraries && <Spinner size="tiny" className={styles.progressSpinner} />}
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
              <Tooltip
                content={this.state.selectedPage ? strings.TooltipProcessPage : strings.TooltipFindBrokenLinks}
                relationship="label">
                <Button icon={<Link24Regular />} onClick={() => this.StartBrokenLinkProcess()} disabled={this.state.isProcessingBrokenLinks}>
                  {!this.state.selectedPage && <span>{strings.FindBrokenLinks}</span>}
                  {this.state.selectedPage && <span>{strings.ProcessPage}</span>}
                </Button>
              </Tooltip>
              {this.state.isProcessingBrokenLinks && <Spinner size="tiny" className={styles.progressSpinner} />}
              &nbsp;
              <Tooltip content={strings.TooltipOpenPageDetails} relationship="label">
                <Button icon={<Open24Regular />} onClick={() => this.ShowPageReport()} disabled={!this.state.selectedPage}>{strings.OpenDetails}</Button>
              </Tooltip>
              &nbsp;
              <Tooltip content={strings.TooltipShowPermissions} relationship="label">
                <Button icon={<KeyMultiple24Regular />} onClick={() => this.ShowPagePermissions()} disabled={!this.state.selectedPage}>{strings.PermissionsButtonLabel}</Button>
              </Tooltip>
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
              <DialogTitle>{strings.PageReportTitle}</DialogTitle>
              <DialogContent style={{ padding: 12 }}>
                {this.state.selectedPage ? (
                  <div>
                    <div><strong>{strings.TitleLabel}</strong> {this.state.selectedPage.title || this.state.selectedPage.name}</div>
                    <div><strong>{strings.UrlLabel}</strong> <a href={this.state.selectedPage.webUrl} target={'_blank'} rel={'noreferrer'}>{this.state.selectedPage.webUrl}</a></div>
                    {(() => {
                      const entry = this.state.pageResults.filter((x: PageResult) => x.pageID === this.state.selectedPage!.id)[0];
                      if (entry) {
                        return (
                          <div style={{ marginTop: 8 }}>
                            <div><strong>{strings.TotalLinksLabel}</strong> {entry.Links.length}</div>
                            <div><strong>{strings.BrokenLinksLabel}</strong> {entry.Links.filter((l: LinkInfo) => l.IsBroken).length}</div>
                            <div style={{ marginTop: 12 }}>
                              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
                                <div><strong>{strings.AllLinksLabel}</strong></div>
                                <Toggle
                                  checked={this.state.showOnlyBrokenLinks}
                                  onChange={(ev, checked?: boolean) => {
                                    this.setState({ showOnlyBrokenLinks: checked || false });
                                  }}
                                  label={strings.ShowOnlyBrokenLinks}
                                  inlineLabel={true}
                                />
                              </div>
                              <div style={{ maxHeight: '300px', overflowY: 'auto', marginTop: 8, border: '1px solid #ccc', padding: 8 }}>
                                {(() => {
                                  const filteredLinks = this.state.showOnlyBrokenLinks 
                                    ? entry.Links.filter((l: LinkInfo) => l.IsBroken)
                                    : entry.Links;
                                  return filteredLinks.length > 0 ? (
                                    filteredLinks.map((link: LinkInfo, index: number) => (
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
                                        <div><strong>{strings.TitleLabel}</strong> {link.title || strings.NoTitle}</div>
                                        <div><strong>{strings.UrlLabel}</strong> 
                                          <a href={link.url} target="_blank" rel="noopener noreferrer" style={{ marginLeft: '4px', color: '#0078d4' }}>
                                          {link.title || strings.NoTitle}
                                          </a>
                                        </div>
                                        {link.Content && link.Content.trim().length > 0 && (
                                          <div style={{ marginTop: '8px' }}>
                                            <button
                                              title={strings.TooltipToggleContent}
                                              onClick={() => {
                                                const currentExpanded = this.state.expandedContentSections || new Set<string>();
                                                const expanded = new Set<string>();
                                                currentExpanded.forEach(val => expanded.add(val));
                                                if (expanded.has(link.url)) {
                                                  expanded.delete(link.url);
                                                } else {
                                                  expanded.add(link.url);
                                                }
                                                this.setState({ expandedContentSections: expanded });
                                              }}
                                              style={{
                                                display: 'flex',
                                                alignItems: 'center',
                                                gap: '4px',
                                                background: 'none',
                                                border: 'none',
                                                cursor: 'pointer',
                                                color: '#0078d4',
                                                padding: '4px 0',
                                                fontSize: '14px'
                                              }}
                                            >
                                              {((this.state.expandedContentSections || new Set<string>()).has(link.url) ? <ChevronUp24Regular /> : <ChevronDown24Regular />)}
                                              <span>{strings.ShowContent}</span>
                                            </button>
                                            {(this.state.expandedContentSections || new Set<string>()).has(link.url) && (
                                              <div
                                                style={{
                                                  marginTop: '8px',
                                                  padding: '8px',
                                                  backgroundColor: '#f9f9f9',
                                                  border: '1px solid #e0e0e0',
                                                  borderRadius: '4px',
                                                  maxHeight: '300px',
                                                  overflowY: 'auto'
                                                }}
                                                dangerouslySetInnerHTML={{ __html: link.Content }}
                                              />
                                            )}
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                    ))
                                  ) : null;
                                })()}
                                {(() => {
                                  const filteredLinks = this.state.showOnlyBrokenLinks 
                                    ? entry.Links.filter((l: LinkInfo) => l.IsBroken)
                                    : entry.Links;
                                  return filteredLinks.length === 0 ? (
                                    <div style={{ padding: '8px', color: '#666', fontStyle: 'italic' }}>
                                      {this.state.showOnlyBrokenLinks 
                                        ? strings.NoBrokenLinksFound
                                        : strings.NoLinksFound}
                                    </div>
                                  ) : null;
                                })()}
                              </div>
                            </div>
                          </div>
                        );
                      }
                      return <div style={{ marginTop: 8 }}>{strings.NoLinkAnalysisAvailable}</div>;
                    })()}
                  </div>
                ) : (
                  <div>{strings.NoItemSelected}</div>
                )}
              </DialogContent>
              <DialogActions>
                <Tooltip content={strings.TooltipCloseDialog} relationship="label">
                  <Button icon={<Dismiss24Regular />} appearance={'secondary'} onClick={() => this.setState({ isReportOpen: false })}>{strings.CloseButton}</Button>
                </Tooltip>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>

        <Dialog open={!!this.state.isLibraryReportOpen} onOpenChange={(_: any, data: any) => this.setState({ isLibraryReportOpen: !!data.open })} modalType={'alert'}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{strings.LibraryReportTitle}</DialogTitle>
              <DialogContent style={{ padding: 12 }}>
                {this.state.selectedLibrary ? (
                  <div>
                    <div><strong>{strings.TitleLabel}</strong> {this.state.selectedLibrary.Title || strings.NA}</div>
                    <div><strong>{strings.TemplateLabel}</strong> {ListTemplateType[this.state.selectedLibrary.BaseTemplate] || strings.NA}</div>
                    <div><strong>{strings.DescriptionLabel}</strong> {this.state.selectedLibrary.Description || strings.NA}</div>
                    <div><strong>{strings.ItemCountLabel}</strong> {this.state.selectedLibrary.ItemCount}</div>
                    <div><strong>{strings.CreatedLabel}</strong> {new Date(this.state.selectedLibrary.Created).toLocaleDateString()}</div>
                    <div><strong>{strings.LastModifiedLabel}</strong> {new Date(this.state.selectedLibrary.LastItemModifiedDate).toLocaleString()}</div>
                    <div><strong>{strings.LastUserModifiedLabel}</strong> {new Date(this.state.selectedLibrary.LastItemUserModifiedDate).toLocaleString()}</div>
                    {this.state.selectedLibrary.LastItemDeletedDate && (
                      <div><strong>{strings.LastDeletedLabel}</strong> {new Date(this.state.selectedLibrary.LastItemDeletedDate).toLocaleString()}</div>
                    )}
                    <div><strong>{strings.EnableVersioningLabel}</strong> {this.state.selectedLibrary.EnableVersioning ? strings.Yes : strings.No}</div>
                    <div><strong>{strings.EnableAttachmentsLabel}</strong> {this.state.selectedLibrary.EnableAttachments ? strings.Yes : strings.No}</div>
                    <div><strong>{strings.EnableFolderCreationLabel}</strong> {this.state.selectedLibrary.EnableFolderCreation ? strings.Yes : strings.No}</div>
                    
                    <div style={{ marginTop: 16 }}>
                      <h4>{strings.OverviewListEntries}</h4>
                      {this.state.selectedLibrary.FoundItems && this.state.selectedLibrary.FoundItems.length > 0 ? (
                        <div>
                          <div><strong>{strings.TotalItemsFound}</strong> {this.state.selectedLibrary.FoundItems.length}</div>
                          <Tooltip content={strings.TooltipShowPermissions} relationship="label">
                            <Button
                              icon={<KeyMultiple24Regular />}
                              onClick={this.onShowPermissionsClick}
                              disabled={!this.state.selectedFoundItem}
                              appearance="secondary"
                              style={{ marginBottom: '8px' }}
                            >
                              {strings.ShowPermissions}
                            </Button>
                          </Tooltip>
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
                          <p style={{ margin: 0, color: '#666' }}>{strings.QueryLibraryForResults}</p>
                        </div>
                      )}
                    </div>
                  </div>
                ) : (
                  <div>{strings.NoLibrarySelected}</div>
                )}
              </DialogContent>
              <DialogActions>
                <Tooltip content={strings.TooltipCloseDialog} relationship="label">
                  <Button icon={<Dismiss24Regular />} appearance={'secondary'} onClick={() => this.setState({ isLibraryReportOpen: false })}>{strings.CloseButton}</Button>
                </Tooltip>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>

        <Dialog open={!!this.state.isPagePermissionsOpen} onOpenChange={(_: any, data: any) => this.setState({ isPagePermissionsOpen: !!data.open })} modalType={'alert'}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{strings.PagePermissionsTitle}</DialogTitle>
              <DialogContent style={{ padding: 12 }}>
                {this.state.selectedPage ? (
                  <div>
                    <div><strong>{strings.TitleLabel}</strong> {this.state.selectedPage.title || this.state.selectedPage.name}</div>
                    <div><strong>{strings.UrlLabel}</strong> <a href={this.state.selectedPage.webUrl} target={'_blank'} rel={'noreferrer'}>{this.state.selectedPage.webUrl}</a></div>
                    {this.state.isLoadingPagePermissions && <Spinner size="tiny" className={styles.progressSpinner} />}
                    {this.state.pagePermissionsError && (
                      <div style={{ color: '#d32f2f', marginTop: 8 }}>{this.state.pagePermissionsError}</div>
                    )}
                    {!this.state.isLoadingPagePermissions && !this.state.pagePermissionsError && this.state.pagePermissions.length === 0 && (
                      <div style={{ marginTop: 8 }}>{strings.NoPermissionsFound}</div>
                    )}
                    <div style={{ marginTop: 12 }}>
                      <ListView
                        items={this.state.pagePermissions}
                        viewFields={this.viewFieldsPermissions}
                        compact={true}
                        selectionMode={SelectionMode.none} />
                    </div>
                  </div>
                ) : (
                  <div>{strings.NoItemSelected}</div>
                )}
              </DialogContent>
              <DialogActions>
                <Tooltip content={strings.TooltipCloseDialog} relationship="label">
                  <Button icon={<Dismiss24Regular />} appearance={'secondary'} onClick={() => this.setState({ isPagePermissionsOpen: false })}>{strings.CloseButton}</Button>
                </Tooltip>
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

  private async ShowPagePermissions(): Promise<void> {
    if (!this.state.selectedPage) {
      return;
    }
    this.setState({ isPagePermissionsOpen: true, isLoadingPagePermissions: true, pagePermissions: [], pagePermissionsError: null });
    try {
      const site = this.GetSelectedSite();
      if (!site) {
        throw new Error('No site is selected.');
      }
      if (!this.state.selectedPage.webUrl) {
        throw new Error('The selected page has no URL.');
      }
      const artefact = await this.permissionsManager.resolveArtefactFromFileUrl(site.url, this.state.selectedPage.webUrl);
      const permissions = await this.permissionsManager.get4ArtefactPermissions(artefact);
      this.setState({ pagePermissions: permissions });
    } catch (error) {
      console.error('Error retrieving page permissions:', error);
      this.setState({ pagePermissionsError: error instanceof Error ? error.message : String(error) });
    } finally {
      this.setState({ isLoadingPagePermissions: false });
    }
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

    this.setState({ isProcessingBrokenLinks: true });

    console.log(`Starting broken link process for site: ${this.state.selectedSiteId}`);
    console.log(`Processing ${this.state.pageEntries.length} pages...`);

    //const dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    const pageAnalyzer = new PageProcessing();
    try {
      // Iterate over all page entries and get their full content
      
      if (this.state.selectedPage) {
        const fullPageContent = await this.dataManager.GetPageContent(this.state.selectedSiteId, this.state.selectedPage.id);
        const resultLinks = await pageAnalyzer.AnalyzePageContent(fullPageContent.canvasLayout!);
        this.state.pageResults.push({
          pageID: this.state.selectedPage.id,
          Links: resultLinks!
        });
      }
      else {
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
      }        
      /*this.setState({
        pageEntries: this.state.pageEntries
      })*/
    } catch (error) {
      console.error('Error during broken link process:', error);
    } finally {
      this.setState({ isProcessingBrokenLinks: false });
    }
  }

  public async CollectItemsFromListAndLibraries():Promise<void>
  {
    const site : Site = this.GetSelectedSite();
    console.log(this.state.selectedLibrary);
    if (this.state.selectedLibrary) {
      const items = await this.dataManager.Query4ItemByDate(
        site,
        this.state.selectedLibrary.Id,        
        this.state.selectedLibrary.ParentWebUrl!,
        this.state.dateStartDate!
      );
      this.state.selectedLibrary.FoundItems = items;
    } 
    else 
    {
      for (const listInfo of this.state.libraryEntries) {
        const items = await this.dataManager.Query4ItemByDate(
          site,
          listInfo.Id,        
          listInfo.ParentWebUrl!,
          this.state.dateStartDate!
        );
        listInfo.FoundItems = items;            
        this.setState({ 
          libraryEntries: this.state.libraryEntries      
        });   
      }
    }
    this.setState({ 
      libraryEntries: this.state.libraryEntries      
    });
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
    await this.GetCheckedOutItems();
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

  private async UpdateLibraryFilter(chkShowLibaries: boolean, chkShowLists: boolean): Promise<void> {
    if (!chkShowLibaries && !chkShowLists) {
      this.setState({ chkShowLibaries, chkShowLists, libraryEntries: [] });
      return;
    }

    this.setState({ isFilteringLibraries: true, chkShowLibaries, chkShowLists });
    try {
      const libraries = await this.dataManager.GetAllLists(this.GetSelectedSite().url, chkShowLists, chkShowLibaries);
      this.setState({ libraryEntries: libraries });
    } finally {
      this.setState({ isFilteringLibraries: false });
    }
  }

  private onDropdDownSelectionChanged = async (event: any, data: any): Promise<void> => {
    const dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    this.setState({ isFilteringLibraries: true });
    const pages = await dataManager.GetPages4Site(data.optionValue);
    this.setState({
      selectedTabValue: this.state.selectedTabValue === null ? "tab1":this.state.selectedTabValue,
      pageEntries: pages,
      selectedSiteId: data.optionValue
    });
    try {
      const siteInfo : Site = this.state.SelectedSites.filter(x=>x.id === data.optionValue)[0];
      const libraries = await dataManager.GetAllLists(siteInfo.url, this.state.chkShowLists, this.state.chkShowLibaries);
      console.log("All lists", libraries);
      this.setState({
        libraryEntries: libraries
      });
    } finally {
      this.setState({ isFilteringLibraries: false });
    }
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
