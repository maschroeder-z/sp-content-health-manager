import * as React from 'react';
import styles from './ContentHealthManager.module.scss';
import type { IContentHealthManagerProps } from './IContentHealthManagerProps';
import { ListView, type IViewField } from '@pnp/spfx-controls-react/lib/ListView';
import { Checkbox, DatePicker, SelectionMode, Toggle } from '@fluentui/react';
import type { IPersonaProps } from '@fluentui/react/lib/Persona';
import { SitePicker } from "@pnp/spfx-controls-react/lib/SitePicker";
import type { Site } from '../../../models/Site';
import { Button, Dropdown, Option, Dialog, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions, Field, TabList, Tab, TabValue, Spinner, Tooltip, Tree, TreeItem, TreeItemLayout, TreeOpenChangeData, TreeOpenChangeEvent } from '@fluentui/react-components';
import { Panel, PanelGroup, PanelResizeHandle } from 'react-resizable-panels';
import { PeoplePicker, IPeoplePickerUserItem, PrincipalType as PickerPrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import GraphDataManager from '../../../services/GraphDataManager';
import { PageProcessing } from '../../../Core/PageProcessing';
import { Page } from '../../../models/Page';
import { PageResult } from '../../../models/PageResult';
import type { LinkInfo } from '../../../models/LinkInfo';
import { CheckmarkCircleColor, CheckmarkCircleHintRegular, FlagPrideIntersexInclusiveProgressFilled, QuestionCircleColor, WarningColor, Search24Regular, DataTrending24Regular, List24Regular, Link24Regular, Clock24Regular, LockClosed24Regular, LockOpen24Regular, ChevronDown24Regular, ChevronUp24Regular, DatabaseSearch24Regular, Open24Regular, Dismiss24Regular, KeyMultiple24Regular, Info24Regular, PeopleTeam16Regular, Person16Regular, DocumentCheckmark24Regular, Library16Regular, List16Regular } from "@fluentui/react-icons";
import { ListInformation } from '../../../models/REST/ListInformation';
import PermissionsManager from '../../../services/PermissionsManager';
import { DirectoryRoleOption, PageStatusInfo, ResolvedGroupUser, SHAREPOINT_RELEVANT_ENTRA_ROLES, SharePointArtefact, SharePointArtefactType, SharePointGroupInfo, SharePointPermissionInfo, SharePointPrincipalPermission } from '../../../models/REST/Permissions';
import { FieldDateRenderer, FieldTextRenderer } from '@pnp/spfx-controls-react';
import { ListTemplateType } from '../../../Core/ListTemplateTypes';
import * as strings from 'ContentHealthManagerWebPartStrings';
//import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta"

interface IPermissionGroupNode {
  key: string;
  groupInfo: SharePointGroupInfo;
  children?: IPermissionGroupNode[];
  isLoadingChildren?: boolean;
  loadError?: string | null;
}

interface IContentHealthManagerState {
  libraryEntries: ListInformation[];
  pageEntries: Page[];
  SelectedSites: Site[];
  selectedSiteId: string | null;
  pageResults: PageResult[];
  isReportOpen?: boolean;
  selectedPage?: Page | null;
  dateStartDate: Date | undefined | null;
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
  permissionGroupTree: IPermissionGroupNode[];
  openTreeNodeKeys: Set<string>;
  selectedTreeNodeKey: string;
  groupMemberCache: Map<string, ResolvedGroupUser[]>;
  isLoadingGroupMembers?: boolean;
  groupMembersError?: string | null;
  currentArtefact: SharePointArtefact | null;
  permissionsSubjectTitle: string;
  permissionsSubjectUrl: string | null;
  isCheckingPrincipalAccess?: boolean;
  principalAccessResult: { displayName: string; hasAccess: boolean; permissionInfo: SharePointPermissionInfo } | null;
  principalAccessError?: string | null;
  pageDetailsCache: Map<string, PageStatusInfo>;
  isLoadingPageDetails?: boolean;
  pageDetailsLoaded?: boolean;
  pageDetailsError?: string | null;
  selectedDirectoryRoleId: string | null;
  directoryRoleMembers: ResolvedGroupUser[];
  isLoadingDirectoryRoleMembers?: boolean;
  directoryRoleMembersError?: string | null;
  permissionsDialogTabValue: TabValue;
}

export default class ContentHealthManager extends React.Component<IContentHealthManagerProps, IContentHealthManagerState> {
  tempSelectedSites: Site[] = [];
  /*
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
   */
  dataManager: GraphDataManager;
  permissionsManager: PermissionsManager;
  // The SitePicker's built-in "clear all" (x) icon only clears its own internal
  // selection state and never invokes the onChange prop, so we detect that click
  // directly in the DOM (capture phase, before the icon's own handler stops
  // propagation) to keep our app state in sync.
  private sitePickerContainerRef: React.RefObject<HTMLDivElement> = React.createRef();
  // Fallback disambiguator for tree node keys when a principal has neither a principalId nor a
  // loginName (a role assignment/group member whose Member expand came back empty) - without this,
  // every such row would collide on the same "login:undefined" key and only the last would render.
  private unresolvedPrincipalCounter = 0;
  // View fields for found items in library report dialog
  viewFieldsFoundItems: IViewField[] = [
    { name: 'Id', displayName: 'ID', sorting: true, isResizable: false, linkPropertyName: 'webUrl' },
    { name: 'Title', displayName: 'Title', sorting: true, isResizable: true },
    {
      name: 'Created', displayName: 'Created', sorting: true, isResizable: false,
      render: (item: any, index, column) => {
        const date = new Date(item.Created);
        return <FieldDateRenderer text={date.toLocaleDateString()} />;
      }
    },
    {
      name: 'Modified', displayName: 'Modified', sorting: true, isResizable: true,
      render: (item: any, index, column) => {
        const date = new Date(item.Modified);
        return <FieldDateRenderer text={date.toLocaleDateString()} />;
      }
    },
    {
      name: 'ContentTypeId', displayName: 'Content Type', sorting: true, isResizable: true,
      render: (item: any, inxdex, column) => {
        if (typeof item.ContentType !== "undefined")
          return item.ContentType;
        return item["ContentType.Name"];
      }
    },
    {
      name: 'CheckedOutBy', displayName: strings.CheckedOutLabel, sorting: true, isResizable: true,
      render: (item: any) => {
        if (this.state.selectedLibrary && !this.SupportsCheckout(this.state.selectedLibrary))
          return <span>{strings.CheckoutNotSupported}</span>;
        return <span>{item.CheckedOutBy || ''}</span>;
      }
    }
  ];

  // BaseTemplate BaseType EnableAttachments EnableFolderCreation EnableVersioning ForceCheckout ItemCount LastItemModifiedDate LastItemUserModifiedDate
  viewFieldsLibs: IViewField[] = [
    {
      name: 'Title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 120,
      render: (item: ListInformation) => {
        // BaseType: 1 = Document Library, everything else (0 = Generic List, etc.) is a list.
        const isLibrary = item.BaseType === 1;
        const TypeIcon = isLibrary ? Library16Regular : List16Regular;
        // ServerRelativeUrl is relative to the tenant root, not the workbench/host origin,
        // so it's resolved against the selected site's own origin rather than used as-is.
        // Falls back to the list settings page (always resolvable from Id) if the default
        // view's URL wasn't returned by the lists REST call for some reason.
        const siteUrl = this.GetSelectedSite()?.url;
        const originMatch = siteUrl?.match(/^https?:\/\/[^/]+/);
        const origin = originMatch ? originMatch[0] : undefined;
        const serverRelativeUrl = item.DefaultView?.ServerRelativeUrl;
        const href = origin
          ? (serverRelativeUrl ? `${origin}${serverRelativeUrl}` : `${siteUrl}/_layouts/15/listedit.aspx?List=${item.Id}`)
          : undefined;
        return (
          <a href={href} target={'_blank'} rel={'noreferrer'} title={isLibrary ? strings.LibraryTypeLabel : strings.ListTypeLabel}>
            <TypeIcon className={styles.inlineIcon} />{item.Title}
          </a>
        );
      }
    },
    { name: 'ItemCount', displayName: 'Items', sorting: true, isResizable: true, minWidth: 120 },
    {
      name: 'FoundItems', displayName: strings.FoundLabel, sorting: true, isResizable: true, minWidth: 120,
      render: (item: ListInformation, index, column) => {
        const entry = this.GetLibraryEntryByIndex(item.Id);
        if (typeof entry.FoundItems !== "undefined" && entry.FoundItems !== null) {
          return <FieldTextRenderer text={`${strings.FoundLabel}: ${entry.FoundItems?.length}`} />;
        }
        else
          return <FieldTextRenderer text={strings.StartQueryForResults} />;
      }
    },
    {
      name: 'Created', displayName: strings.CreatedAtLabel, sorting: true, isResizable: true, minWidth: 100,
      render: (item: ListInformation, index, column) => {
        const date = new Date(item.Created);
        return <FieldDateRenderer text={date.toLocaleDateString()} />;
      }
    },
    {
      name: 'LastItemModifiedDate', displayName: strings.LastChangeLabel, sorting: true, isResizable: true, minWidth: 120, linkPropertyName: 'webUrl',
      render: (item: ListInformation, index, column) => {
        const date = new Date(item.LastItemModifiedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;
      }
    },
    {
      name: 'LastItemUserModifiedDate', displayName: strings.UserChangedLabel, sorting: true, isResizable: true, minWidth: 120, linkPropertyName: 'webUrl',
      render: (item: ListInformation, index, column) => {
        const date = new Date(item.LastItemUserModifiedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;
      }
    },
    {
      name: 'LastItemDeletedDate', displayName: strings.LastDeletionLabel, sorting: true, isResizable: true, minWidth: 100,
      render: (item: ListInformation, index, column) => {
        const date = new Date(item.LastItemDeletedDate);
        return <FieldDateRenderer text={date.toLocaleString()} />;
      }
    },
    { name: 'ItemCount', displayName: 'Items', sorting: true, isResizable: true, minWidth: 120 },
    {
      name: 'FoundItems', displayName: strings.FoundLabel, sorting: true, isResizable: true, minWidth: 120,
      render: (item: ListInformation, index, column) => {
        const entry = this.GetLibraryEntryByIndex(item.Id);
        if (entry.FoundItemsUnsupported) {
          return <FieldTextRenderer text={strings.CheckoutNotSupported} />;
        }
        else if (typeof entry.FoundCheckedOutItems !== "undefined" && entry.FoundCheckedOutItems !== null) {
          return <FieldTextRenderer text={`${strings.FoundLabel}: ${entry.FoundCheckedOutItems?.length}`} />;
        }
        else
          return <FieldTextRenderer text={strings.StartQueryForResults} />;
      }
    },
    { name: 'Description', displayName: 'Description', sorting: true, isResizable: true, minWidth: 100 }
  ];

  viewFieldsPage: IViewField[] = [
    { name: 'title', displayName: 'Title', sorting: true, isResizable: true, minWidth: 50, linkPropertyName: 'webUrl' },
    { name: 'name', displayName: 'Name', sorting: true, isResizable: true, minWidth: 200 },
    {
      name: 'Links', displayName: 'Links', sorting: false, isResizable: true,
      render: (item, index, column) => {
        const entry = this.state.pageResults.filter(x => x.pageID === item.id)[0];

        if (typeof entry === "undefined" || typeof entry.Links === "undefined") {
          return <>
            <CheckmarkCircleHintRegular />
          </>;
        }

        if (entry.Links.filter(x => x.IsBroken).length > 0) {
          return (<>
            <WarningColor />
            &nbsp;<span>{strings.FoundLinksCount.replace('{0}', entry.Links.length.toString()).replace('{1}', entry.Links.filter(x => x.IsBroken).length.toString())}</span>
          </>);
        }
        return <>
          <CheckmarkCircleColor />
          &nbsp;
          <span>{strings.FoundLinksCount.replace('{0}', entry.Links.length.toString()).replace('{1}', entry.Links.filter(x => x.IsBroken).length.toString())}</span>
        </>;
      }
    }
  ];

  private getPageViewFields(): IViewField[] {
    if (!this.state.pageDetailsLoaded) {
      return this.viewFieldsPage;
    }
    return [
      ...this.viewFieldsPage,
      {
        name: 'needsApproval', displayName: strings.NeedsApprovalLabel, sorting: false, isResizable: true, minWidth: 140,
        render: (item: Page) => {
          const status = this.state.pageDetailsCache.get(item.id);
          if (!status) {
            return <></>;
          }
          return (
            <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
              {status.needsApproval ? <WarningColor /> : <CheckmarkCircleColor />}
              <span>{status.needsApproval ? strings.Yes : strings.No}</span>
            </span>
          );
        }
      },
      {
        name: 'hasUniquePermission', displayName: strings.HasUniquePermissionLabel, sorting: false, isResizable: true, minWidth: 160,
        render: (item: Page) => {
          const status = this.state.pageDetailsCache.get(item.id);
          if (!status) {
            return <></>;
          }
          return (
            <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
              {status.hasUniquePermission ? <LockClosed24Regular /> : <LockOpen24Regular />}
              <span>{status.hasUniquePermission ? strings.Yes : strings.No}</span>
            </span>
          );
        }
      },
      {
        name: 'checkedOutBy', displayName: strings.CheckedOutLabel, sorting: false, isResizable: true, minWidth: 160,
        render: (item: Page) => {
          const status = this.state.pageDetailsCache.get(item.id);
          if (!status) {
            return <></>;
          }
          return (
            <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
              {status.checkedOutBy ? <Person16Regular /> : <CheckmarkCircleColor />}
              <span>{status.checkedOutBy || strings.NotCheckedOut}</span>
            </span>
          );
        }
      }
    ];
  }

  viewFieldsPermissions: IViewField[] = [
    { name: 'displayName', displayName: strings.PrincipalNameLabel, sorting: true, isResizable: true, minWidth: 180 },
    {
      name: 'isGroup', displayName: strings.PrincipalTypeLabel, sorting: true, isResizable: true, minWidth: 100,
      render: (item: SharePointPrincipalPermission) => (
        <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
          {item.isGroup ? <PeopleTeam16Regular /> : <Person16Regular />}
          <span>{item.isGroup ? strings.GroupLabel : strings.UserLabel}</span>
        </span>
      )
    },
    {
      name: 'loginName', displayName: strings.LoginNameLabel, sorting: false, isResizable: true, minWidth: 220,
      render: (item: SharePointPrincipalPermission) => <span title={item.loginName}>{this.formatLoginName(item.loginName)}</span>
    },
    {
      name: 'roles', displayName: strings.RolesLabel, sorting: false, isResizable: true, minWidth: 200,
      render: (item: SharePointPrincipalPermission) => <FieldTextRenderer text={(item.roles || []).join(', ')} />
    }
  ];

  viewFieldsGroupMembers: IViewField[] = [
    { name: 'displayName', displayName: strings.PrincipalNameLabel, sorting: true, isResizable: true, minWidth: 180 },
    { name: 'email', displayName: strings.EmailLabel, sorting: true, isResizable: true, minWidth: 220 },
    {
      name: 'loginName', displayName: strings.LoginNameLabel, sorting: false, isResizable: true, minWidth: 220,
      render: (item: ResolvedGroupUser) => <span title={item.loginName}>{this.formatLoginName(item.loginName)}</span>
    }
  ];

  // Claims-encoded login names look like "i:0#.f|membership|user@tenant.com" or
  // "c:0t.c|tenant|<aadGroupId>" - strip the claims provider prefix and keep the
  // human-meaningful part (email/UPN or the trailing id) for display.
  private formatLoginName(loginName: string | undefined): string {
    if (!loginName) {
      return '';
    }
    const lastSegment = loginName.split('|').pop();
    return lastSegment || loginName;
  }

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
      pagePermissionsError: null,
      permissionGroupTree: [],
      openTreeNodeKeys: new Set<string>(),
      selectedTreeNodeKey: 'root',
      groupMemberCache: new Map<string, ResolvedGroupUser[]>(),
      isLoadingGroupMembers: false,
      groupMembersError: null,
      currentArtefact: null,
      permissionsSubjectTitle: '',
      permissionsSubjectUrl: null,
      isCheckingPrincipalAccess: false,
      principalAccessResult: null,
      principalAccessError: null,
      pageDetailsCache: new Map<string, PageStatusInfo>(),
      isLoadingPageDetails: false,
      pageDetailsLoaded: false,
      pageDetailsError: null,
      selectedDirectoryRoleId: null,
      directoryRoleMembers: [],
      isLoadingDirectoryRoleMembers: false,
      directoryRoleMembersError: null,
      permissionsDialogTabValue: 'permissions'
    };
    this.dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    this.permissionsManager = new PermissionsManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
  }

  private GetLibraryEntryByIndex(index: string): ListInformation {
    return this.state.libraryEntries.filter(x => x.Id === index)[0];
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
                    <li><DocumentCheckmark24Regular className={styles.inlineIcon} />{strings.PagesWaitingForApproval}</li>
                  </ul>
                </li>
                <li>
                  <KeyMultiple24Regular className={styles.inlineIcon} />
                  <strong>{strings.FourthCheckPermissions.split(' - ')[0]}</strong> - {strings.FourthCheckPermissions.split(' - ')[1]}
                </li>
              </ol>
            </div>
          </div>
        )}

        <div className={styles.row}>
          <div className={styles['col-sm12']}>
            {this.state.SelectedSites.length === 0 && <p className={styles.infoMessage}><QuestionCircleColor />{strings.SelectFirstAllSites}</p>}
            <Field label={strings.SelectSitesLabel}>
              <div ref={this.sitePickerContainerRef}>
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
                    const newSites = (sites || []) as Site[];
                    const evaluatedSiteRemoved = this.state.selectedSiteId !== null
                      && !newSites.some(s => s.id === this.state.selectedSiteId);
                    if (newSites.length === 0 || evaluatedSiteRemoved) {
                      this.resetAppState(newSites);
                    } else {
                      this.setState({ SelectedSites: newSites });
                    }
                  }}
                  placeholder={strings.SelectAllSitesPlaceholder}
                  searchPlaceholder={strings.FilterSitesPlaceholder} />
              </div>
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
                  {this.state.SelectedSites.map((entry: Site) => (
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
          </TabList> </>}

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
                    minDate={new Date(2000, 0, 1)}
                    maxDate={new Date()}
                    placeholder={strings.SelectQueryDatePlaceholder}
                    onSelectDate={(selectedDate: Date | undefined | null) => this.setState(
                      { dateStartDate: selectedDate }
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
            <div className={`${styles.row} ${styles.libraryCommands} ${styles.libraryActionsRow}`}>
              <div className={styles.libraryActionsButtons}>
                <Tooltip content={strings.TooltipOpenLibraryDetails} relationship="label">
                  <Button icon={<Open24Regular />} onClick={() => this.ShowLibraryReport()} disabled={!this.state.selectedLibrary}>{strings.OpenDetails}</Button>
                </Tooltip>
                <Tooltip content={strings.TooltipShowSelectedLibraryPermissions} relationship="label">
                  <Button icon={<KeyMultiple24Regular />} onClick={() => this.ShowPagePermissions()} disabled={!this.state.selectedLibrary}>{strings.PermissionsButtonLabel}</Button>
                </Tooltip>
              </div>

              <div className={styles.checkboxContainer}>
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
                <Tooltip content={this.state.selectedPage ? strings.TooltipShowPermissions : strings.TooltipShowLibraryPermissions} relationship="label">
                  <Button icon={<KeyMultiple24Regular />} onClick={() => this.ShowPagePermissions()}>{strings.PermissionsButtonLabel}</Button>
                </Tooltip>
                &nbsp;
                <Tooltip content={strings.TooltipLoadPageDetails} relationship="label">
                  <Button
                    icon={<Info24Regular />}
                    onClick={() => this.LoadPageDetails()}
                    disabled={this.state.isLoadingPageDetails || this.state.pageDetailsLoaded || this.state.pageEntries.length === 0}>
                    {strings.LoadPageDetailsButtonLabel}
                  </Button>
                </Tooltip>
                {this.state.isLoadingPageDetails && <Spinner size="tiny" className={styles.progressSpinner} />}
                {this.state.pageDetailsError && (
                  <div style={{ color: '#d32f2f' }}>{this.state.pageDetailsError}</div>
                )}
              </div>
            </div>
            <ListView
              items={this.state.pageEntries}
              viewFields={this.getPageViewFields()}
              compact={true}
              selectionMode={SelectionMode.single}
              selection={this.onListSelectionChanged} />
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
                      {(this.state.selectedLibrary.FoundItems && this.state.selectedLibrary.FoundItems.length > 0)
                        || ((this.state.selectedLibrary.FoundCheckedOutItems && this.state.selectedLibrary.FoundCheckedOutItems.length > 0)) ? (
                        <div>
                          {this.state.selectedLibrary.FoundItems && this.state.selectedLibrary.FoundItems.length > 0 ? (
                            <>
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
                            </>) : null}
                          {this.state.selectedLibrary.FoundCheckedOutItems && this.state.selectedLibrary.FoundCheckedOutItems.length > 0 ? (
                            <>
                              <div><strong>{strings.TotalCheckedOutIemsFound}</strong> {this.state.selectedLibrary.FoundCheckedOutItems.length}</div>
                              <div style={{ marginTop: 8, maxHeight: '300px' }}>
                                <ListView
                                  items={this.state.selectedLibrary.FoundCheckedOutItems}
                                  viewFields={this.viewFieldsFoundItems}
                                  compact={true}
                                  selectionMode={SelectionMode.single}
                                  selection={this.onFoundItemSelectionChanged}
                                />
                              </div>
                            </>) : null}
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
          <DialogSurface style={{ maxWidth: '95vw', width: 960 }}>
            <DialogBody>
              <DialogTitle>{strings.PagePermissionsTitle}</DialogTitle>
              <DialogContent style={{ padding: 12 }}>
                <div>
                  <div><strong>{strings.TitleLabel}</strong> {this.state.permissionsSubjectTitle}</div>
                  {this.state.permissionsSubjectUrl && (
                    <div><strong>{strings.UrlLabel}</strong> <a href={this.state.permissionsSubjectUrl} target={'_blank'} rel={'noreferrer'}>{this.state.permissionsSubjectUrl}</a></div>
                  )}
                  <TabList selectedValue={this.state.permissionsDialogTabValue} onTabSelect={this.onPermissionsDialogTabSelect} style={{ marginTop: 12 }}>
                    <Tab value="permissions">{strings.PermissionsDialogPermissionsTab}</Tab>
                    <Tab value="entraRoles">{strings.PermissionsDialogEntraRolesTab}</Tab>
                  </TabList>
                  {this.state.permissionsDialogTabValue === 'permissions' && (
                    <div style={{ marginTop: 12 }}>
                      {this.state.currentArtefact && (
                        <div style={{ marginBottom: 12 }}>
                          <PeoplePicker
                            context={{
                              absoluteUrl: this.state.currentArtefact.webUrl,
                              msGraphClientFactory: this.props.msGraphClientFactory,
                              spHttpClient: this.props.spHTTPClient
                            }}
                            showtooltip={true}
                            personSelectionLimit={1}
                            principalTypes={[PickerPrincipalType.User, PickerPrincipalType.SecurityGroup, PickerPrincipalType.SharePointGroup, PickerPrincipalType.DistributionList]}
                            useSubstrateSearch={false}
                            searchTextLimit={2}
                            placeholder={strings.SearchUserOrGroupPlaceholder}
                            onChange={(items: IPersonaProps[]) => {
                              const item = items && items[0] ? items[0] as unknown as IPeoplePickerUserItem : undefined;
                              if (item) {
                                void this.checkPrincipalAccess(item);
                              } else {
                                this.setState({ principalAccessResult: null, principalAccessError: null });
                              }
                            }}
                          />
                          {this.state.isCheckingPrincipalAccess && <Spinner size="tiny" className={styles.progressSpinner} />}
                          {this.state.principalAccessError && (
                            <div style={{ color: '#d32f2f', marginTop: 8 }}>{this.state.principalAccessError}</div>
                          )}
                          {!this.state.isCheckingPrincipalAccess && !this.state.principalAccessError && this.state.principalAccessResult && (
                            <div style={{ marginTop: 8 }}>
                              {this.state.principalAccessResult.hasAccess
                                ? strings.HasAccessLabel
                                  .replace('{0}', this.state.principalAccessResult.displayName)
                                  .replace('{1}', this.getPermissionLevelLabel(this.state.principalAccessResult.permissionInfo))
                                : strings.NoAccessLabel.replace('{0}', this.state.principalAccessResult.displayName)}
                            </div>
                          )}
                        </div>
                      )}
                      {this.state.isLoadingPagePermissions && <Spinner size="tiny" className={styles.progressSpinner} />}
                      {this.state.pagePermissionsError && (
                        <div style={{ color: '#d32f2f', marginTop: 8 }}>{this.state.pagePermissionsError}</div>
                      )}
                      {!this.state.isLoadingPagePermissions && !this.state.pagePermissionsError && this.state.pagePermissions.length === 0 && (
                        <div style={{ marginTop: 8 }}>{strings.NoPermissionsFound}</div>
                      )}
                      {!this.state.isLoadingPagePermissions && !this.state.pagePermissionsError && this.state.pagePermissions.length > 0 && (
                        <PanelGroup direction="horizontal" style={{ height: 420, marginTop: 12 }}>
                          <Panel defaultSize={30} minSize={15} maxSize={60}>
                            <div style={{ height: '100%', overflow: 'auto', borderRight: '1px solid #e0e0e0' }}>
                              <Tree
                                openItems={this.state.openTreeNodeKeys}
                                onOpenChange={this.handleTreeOpenChange}
                                aria-label={strings.PagePermissionsTitle}
                              >
                                <TreeItem itemType="leaf" value="root">
                                  <TreeItemLayout
                                    onClick={() => this.selectTreeNode('root')}
                                    style={this.state.selectedTreeNodeKey === 'root' ? { background: '#e0e0e0' } : undefined}
                                  >
                                    {this.state.permissionsSubjectTitle}
                                  </TreeItemLayout>
                                </TreeItem>
                                {this.state.permissionGroupTree.map(node => this.renderGroupTreeNode(node))}
                              </Tree>
                            </div>
                          </Panel>
                          <PanelResizeHandle style={{ width: 6, cursor: 'col-resize', background: '#e0e0e0' }} />
                          <Panel>
                            <div style={{ height: '100%', overflow: 'auto', paddingLeft: 8 }}>
                              {this.state.selectedTreeNodeKey === 'root' ? (
                                <ListView
                                  items={this.state.pagePermissions.filter(p => !p.isGroup)}
                                  viewFields={this.viewFieldsPermissions}
                                  compact={true}
                                  selectionMode={SelectionMode.none} />
                              ) : (
                                <>
                                  {this.state.isLoadingGroupMembers && <Spinner size="tiny" className={styles.progressSpinner} />}
                                  {this.state.groupMembersError && (
                                    <div style={{ color: '#d32f2f', marginTop: 8 }}>{this.state.groupMembersError}</div>
                                  )}
                                  {!this.state.isLoadingGroupMembers && !this.state.groupMembersError && (
                                    <ListView
                                      items={this.state.groupMemberCache.get(this.state.selectedTreeNodeKey) || []}
                                      viewFields={this.viewFieldsGroupMembers}
                                      compact={true}
                                      selectionMode={SelectionMode.none} />
                                  )}
                                </>
                              )}
                            </div>
                          </Panel>
                        </PanelGroup>
                      )}
                    </div>
                  )}
                  {this.state.permissionsDialogTabValue === 'entraRoles' && (
                    <div style={{ marginTop: 12 }}>
                      <Field label={strings.DirectoryRolePickerLabel} hint={strings.DirectoryRolePickerHint}>
                        <Dropdown
                          placeholder={strings.SelectDirectoryRolePlaceholder}
                          value={SHAREPOINT_RELEVANT_ENTRA_ROLES.find((r: DirectoryRoleOption) => r.roleTemplateId === this.state.selectedDirectoryRoleId)?.displayName || ''}
                          selectedOptions={this.state.selectedDirectoryRoleId ? [this.state.selectedDirectoryRoleId] : []}
                          onOptionSelect={(_: any, data: any) => data.optionValue && this.selectDirectoryRole(data.optionValue)}
                        >
                          {SHAREPOINT_RELEVANT_ENTRA_ROLES.map((role: DirectoryRoleOption) => (
                            <Option key={role.roleTemplateId} value={role.roleTemplateId}>{role.displayName}</Option>
                          ))}
                        </Dropdown>
                      </Field>
                      {this.state.isLoadingDirectoryRoleMembers && <Spinner size="tiny" className={styles.progressSpinner} />}
                      {this.state.directoryRoleMembersError && (
                        <div style={{ color: '#d32f2f', marginTop: 8 }}>{this.state.directoryRoleMembersError}</div>
                      )}
                      {!this.state.isLoadingDirectoryRoleMembers && !this.state.directoryRoleMembersError && this.state.selectedDirectoryRoleId && (
                        <ListView
                          items={this.state.directoryRoleMembers}
                          viewFields={this.viewFieldsGroupMembers}
                          compact={true}
                          selectionMode={SelectionMode.none} />
                      )}
                    </div>
                  )}
                </div>
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
    this.sitePickerContainerRef.current?.addEventListener('click', this.handleSitePickerClearAllClick, true);
  }

  public componentWillUnmount(): void {
    this.sitePickerContainerRef.current?.removeEventListener('click', this.handleSitePickerClearAllClick, true);
  }

  private handleSitePickerClearAllClick = (event: MouseEvent): void => {
    const target = event.target as HTMLElement | null;
    if (target?.closest('[data-icon-name="Cancel"]')) {
      this.resetAppState([]);
    }
  }

  private ShowLibraryReport(): void {
    if (!this.state.selectedLibrary) {
      return;
    }
    this.setState({ isLibraryReportOpen: true });
  }

  private ShowPageReport(): void {
    if (!this.state.selectedPage) {
      return;
    }
    this.setState({ isReportOpen: true });
  }

  private async ShowPagePermissions(): Promise<void> {
    const site = this.GetSelectedSite();
    if (!site) {
      console.warn('No site selected. Please select a site first.');
      return;
    }
    const isLibraryMode = this.state.selectedTabValue === 'tab2' && !!this.state.selectedLibrary;
    this.setState({
      isPagePermissionsOpen: true,
      permissionsDialogTabValue: 'permissions',
      isLoadingPagePermissions: true,
      pagePermissions: [],
      pagePermissionsError: null,
      permissionGroupTree: [],
      openTreeNodeKeys: new Set<string>(),
      selectedTreeNodeKey: 'root',
      groupMemberCache: new Map<string, ResolvedGroupUser[]>(),
      groupMembersError: null,
      currentArtefact: null,
      permissionsSubjectTitle: isLibraryMode
        ? (this.state.selectedLibrary!.Title || '')
        : this.state.selectedPage ? (this.state.selectedPage.title || this.state.selectedPage.name || '') : strings.PagesLibraryLabel,
      permissionsSubjectUrl: isLibraryMode
        ? this.state.selectedLibrary!.DefaultView.ServerRelativeUrl
        : (this.state.selectedPage?.webUrl || null),
      isCheckingPrincipalAccess: false,
      principalAccessResult: null,
      principalAccessError: null
    });
    try {
      let artefact: SharePointArtefact;
      if (isLibraryMode) {
        artefact = {
          // ListInformation.ParentWebUrl is not a usable web URL (GraphDataManager appends the list's
          // EntityTypeName onto it for a different purpose) - libraryEntries is always fetched for the
          // currently selected site, so that site's own URL is the correct owning web.
          type: SharePointArtefactType.List,
          webUrl: site.url,
          listId: this.state.selectedLibrary!.Id
        };
      } else if (this.state.selectedPage) {
        if (!this.state.selectedPage.webUrl) {
          throw new Error('The selected page has no URL.');
        }
        artefact = await this.permissionsManager.resolveArtefactFromFileUrl(site.url, this.state.selectedPage.webUrl);
      } else {
        artefact = await this.permissionsManager.resolvePagesLibraryArtefact(site.url);
      }
      const permissions = await this.permissionsManager.get4ArtefactPermissions(artefact);
      const groupTree = permissions.filter(p => p.isGroup).map(p => this.buildGroupNode(p, artefact.webUrl));
      this.setState({ pagePermissions: permissions, permissionGroupTree: groupTree, currentArtefact: artefact });
    } catch (error) {
      console.error('Error retrieving page permissions:', error);
      this.setState({ pagePermissionsError: error instanceof Error ? error.message : String(error) });
    } finally {
      this.setState({ isLoadingPagePermissions: false });
    }
  }

  private async checkPrincipalAccess(item: IPeoplePickerUserItem): Promise<void> {
    if (!this.state.currentArtefact) {
      return;
    }
    this.setState({ isCheckingPrincipalAccess: true, principalAccessResult: null, principalAccessError: null });
    try {
      const report = await this.permissionsManager.checkAccess4Principal({ id: item.id, displayName: item.text }, this.state.currentArtefact);
      this.setState({ principalAccessResult: { displayName: item.text, hasAccess: report.hasAccess, permissionInfo: report.permissionInfo } });
    } catch (error) {
      console.error('Error checking principal access:', error);
      this.setState({ principalAccessError: error instanceof Error ? error.message : String(error) });
    } finally {
      this.setState({ isCheckingPrincipalAccess: false });
    }
  }

  private async LoadPageDetails(): Promise<void> {
    const site = this.GetSelectedSite();
    if (!site) {
      return;
    }
    this.setState({ isLoadingPageDetails: true, pageDetailsError: null });
    try {
      const entries = await Promise.all(this.state.pageEntries.map(async (page): Promise<[string, PageStatusInfo]> => {
        const status = await this.permissionsManager.getPageStatus(site.url, page.webUrl!);
        return [page.id, status];
      }));
      this.setState({ pageDetailsCache: new Map(entries), pageDetailsLoaded: true });
    } catch (error) {
      console.error('Error loading page details:', error);
      this.setState({ pageDetailsError: error instanceof Error ? error.message : String(error) });
    } finally {
      this.setState({ isLoadingPageDetails: false });
    }
  }

  private getPermissionLevelLabel(info: SharePointPermissionInfo): string {
    if (info.hasFullControl || info.canManagePermissions) {
      return strings.FullControlLabel;
    }
    if (info.canManageLists) {
      return strings.DesignLabel;
    }
    if (info.canEdit) {
      return strings.EditLabel;
    }
    if (info.canContribute) {
      return strings.ContributeLabel;
    }
    if (info.canView) {
      return strings.ReadLabel;
    }
    return strings.NoAccessLevelLabel;
  }

  private buildGroupNode(source: SharePointPrincipalPermission | SharePointGroupInfo, webUrl: string): IPermissionGroupNode {
    const groupInfo: SharePointGroupInfo = 'webUrl' in source
      ? source
      : {
        webUrl,
        principalId: source.principalId,
        principalType: source.principalType,
        loginName: source.loginName,
        displayName: source.displayName
      };
    const key = groupInfo.principalId !== undefined
      ? `id:${groupInfo.principalId}`
      : groupInfo.loginName
        ? `login:${groupInfo.loginName}`
        : `unresolved:${this.unresolvedPrincipalCounter++}`;
    return { key, groupInfo, children: undefined };
  }

  private findTreeNode(nodes: IPermissionGroupNode[], key: string): IPermissionGroupNode | undefined {
    for (const node of nodes) {
      if (node.key === key) {
        return node;
      }
      if (node.children) {
        const found = this.findTreeNode(node.children, key);
        if (found) {
          return found;
        }
      }
    }
    return undefined;
  }

  private updateTreeNode(nodes: IPermissionGroupNode[], key: string, patch: Partial<IPermissionGroupNode>): IPermissionGroupNode[] {
    return nodes.map(node => {
      if (node.key === key) {
        return { ...node, ...patch };
      }
      if (node.children) {
        return { ...node, children: this.updateTreeNode(node.children, key, patch) };
      }
      return node;
    });
  }

  private handleTreeOpenChange = (_event: TreeOpenChangeEvent, data: TreeOpenChangeData): void => {
    const openKeys = data.openItems as Set<string>;
    this.setState({ openTreeNodeKeys: openKeys });

    if (data.open) {
      const node = this.findTreeNode(this.state.permissionGroupTree, String(data.value));
      if (node && node.children === undefined && !node.isLoadingChildren) {
        void this.loadNestedGroups(node);
      }
    }
  }

  private async loadNestedGroups(node: IPermissionGroupNode): Promise<void> {
    this.setState({ permissionGroupTree: this.updateTreeNode(this.state.permissionGroupTree, node.key, { isLoadingChildren: true, loadError: null }) });
    try {
      const nestedGroups = await this.permissionsManager.resolveNestedGroups(node.groupInfo);
      const children = nestedGroups.map(g => this.buildGroupNode(g, node.groupInfo.webUrl));
      this.setState({ permissionGroupTree: this.updateTreeNode(this.state.permissionGroupTree, node.key, { children, isLoadingChildren: false }) });
    } catch (error) {
      console.error('Error resolving nested groups:', error);
      const message = error instanceof Error ? error.message : String(error);
      this.setState({ permissionGroupTree: this.updateTreeNode(this.state.permissionGroupTree, node.key, { isLoadingChildren: false, loadError: message }) });
    }
  }

  private selectTreeNode(key: string, groupInfo?: SharePointGroupInfo): void {
    this.setState({ selectedTreeNodeKey: key });
    if (key === 'root' || !groupInfo) {
      return;
    }
    if (this.state.groupMemberCache.has(key)) {
      return;
    }
    void this.loadGroupMembers(key, groupInfo);
  }

  private async loadGroupMembers(key: string, groupInfo: SharePointGroupInfo): Promise<void> {
    this.setState({ isLoadingGroupMembers: true, groupMembersError: null });
    try {
      const users = await this.permissionsManager.resolveUser4Group(groupInfo);
      const cache = new Map(this.state.groupMemberCache);
      cache.set(key, users);
      this.setState({ groupMemberCache: cache });
    } catch (error) {
      console.error('Error resolving group members:', error);
      this.setState({ groupMembersError: error instanceof Error ? error.message : String(error) });
    } finally {
      this.setState({ isLoadingGroupMembers: false });
    }
  }

  private selectDirectoryRole(roleTemplateId: string): void {
    this.setState({
      selectedDirectoryRoleId: roleTemplateId,
      directoryRoleMembers: [],
      directoryRoleMembersError: null
    });
    void this.loadDirectoryRoleMembers(roleTemplateId);
  }

  private async loadDirectoryRoleMembers(roleTemplateId: string): Promise<void> {
    this.setState({ isLoadingDirectoryRoleMembers: true, directoryRoleMembersError: null });
    try {
      const members = await this.permissionsManager.resolveDirectoryRoleUsers(roleTemplateId);
      this.setState({ directoryRoleMembers: members });
    } catch (error) {
      console.error('Error resolving directory role members:', error);
      this.setState({ directoryRoleMembersError: error instanceof Error ? error.message : String(error) });
    } finally {
      this.setState({ isLoadingDirectoryRoleMembers: false });
    }
  }

  private renderGroupTreeNode(node: IPermissionGroupNode): JSX.Element {
    return (
      <TreeItem itemType="branch" value={node.key} key={node.key}>
        <TreeItemLayout
          onClick={() => this.selectTreeNode(node.key, node.groupInfo)}
          style={this.state.selectedTreeNodeKey === node.key ? { background: '#e0e0e0' } : undefined}
          iconBefore={<PeopleTeam16Regular />}
        >
          {node.groupInfo.displayName}
        </TreeItemLayout>
        <Tree>
          {node.isLoadingChildren && (
            <TreeItem itemType="leaf" value={`${node.key}-loading`}>
              <TreeItemLayout><Spinner size="tiny" /></TreeItemLayout>
            </TreeItem>
          )}
          {node.loadError && (
            <TreeItem itemType="leaf" value={`${node.key}-error`}>
              <TreeItemLayout><span style={{ color: '#d32f2f' }}>{node.loadError}</span></TreeItemLayout>
            </TreeItem>
          )}
          {node.children?.length === 0 && (
            <TreeItem itemType="leaf" value={`${node.key}-empty`}>
              <TreeItemLayout>{strings.NoNestedGroups}</TreeItemLayout>
            </TreeItem>
          )}
          {node.children?.map(child => this.renderGroupTreeNode(child))}
        </Tree>
      </TreeItem>
    );
  }

  private async StartBrokenLinkProcess(): Promise<void> {
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

  public async CollectItemsFromListAndLibraries(): Promise<void> {
    const site: Site = this.GetSelectedSite();
    console.log(this.state.selectedLibrary);
    if (this.state.selectedLibrary) {
      const items = await this.dataManager.Query4ItemByDate(
        site,
        this.state.selectedLibrary.Id,
        this.state.selectedLibrary.ParentWebUrl!,
        this.state.dateStartDate!
      );
      this.state.selectedLibrary.FoundItems = items;
      //this.state.selectedLibrary.FoundItemsUnsupported = false;
    }
    else {
      for (const listInfo of this.state.libraryEntries) {
        const items = await this.dataManager.Query4ItemByDate(
          site,
          listInfo.Id,
          listInfo.ParentWebUrl!,
          this.state.dateStartDate!
        );
        listInfo.FoundItems = items;
        //listInfo.FoundItemsUnsupported = false;
        this.setState({
          libraryEntries: this.state.libraryEntries
        });
      }
    }
    this.setState({
      libraryEntries: this.state.libraryEntries
    });
  }

  public async GetCheckedOutItems(): Promise<void> {
    const site: Site = this.GetSelectedSite();
    for (const listInfo of this.state.libraryEntries) {
      // Skip lists/libraries that don't support check-out - the "Checked out" column renders
      // a "not supported" message for those instead.
      if (!this.SupportsCheckout(listInfo)) {
        listInfo.FoundCheckedOutItems = [];
        listInfo.FoundItemsUnsupported = true;
        this.setState({
          libraryEntries: this.state.libraryEntries
        });
        continue;
      }
      const items = await this.dataManager.Query4CheckedOutItems(
        site,
        listInfo.Id,
        listInfo.DefaultView.ServerRelativeUrl,
        this.state.dateStartDate!
      );
      listInfo.FoundCheckedOutItems = items;
      listInfo.FoundItemsUnsupported = false;
      this.setState({
        libraryEntries: this.state.libraryEntries
      });
    }
  }

  private async StartQueryCheckedOutItems(): Promise<void> {
    await this.GetCheckedOutItems();
  }

  public async GetPermission4SelectedItem(site: Site, listID: string, listItemID: string): Promise<void> {
    // Mirrors ShowPagePermissions' state contract (isPagePermissionsOpen + pagePermissions +
    // permissionGroupTree) so the found item reuses the same "Page permissions" dialog instead
    // of only logging to the console.
    const item = this.state.selectedFoundItem;
    this.setState({
      isPagePermissionsOpen: true,
      permissionsDialogTabValue: 'permissions',
      isLoadingPagePermissions: true,
      pagePermissions: [],
      pagePermissionsError: null,
      permissionGroupTree: [],
      openTreeNodeKeys: new Set<string>(),
      selectedTreeNodeKey: 'root',
      groupMemberCache: new Map<string, ResolvedGroupUser[]>(),
      groupMembersError: null,
      currentArtefact: null,
      permissionsSubjectTitle: item?.Title || item?.FileLeafRef || '',
      permissionsSubjectUrl: item?.webUrl || null,
      isCheckingPrincipalAccess: false,
      principalAccessResult: null,
      principalAccessError: null
    });
    try {
      const artefact: SharePointArtefact = {
        type: SharePointArtefactType.ListItem,
        webUrl: site.url,
        listId: listID,
        itemId: Number(listItemID)
      };
      const permissions = await this.permissionsManager.get4ArtefactPermissions(artefact);
      const groupTree = permissions.filter(p => p.isGroup).map(p => this.buildGroupNode(p, artefact.webUrl));
      this.setState({ pagePermissions: permissions, permissionGroupTree: groupTree, currentArtefact: artefact });
    } catch (error) {
      console.error('Error retrieving item permissions:', error);
      this.setState({ pagePermissionsError: error instanceof Error ? error.message : String(error) });
    } finally {
      this.setState({ isLoadingPagePermissions: false });
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

  private resetAppState(sites: Site[] = []): void {
    this.resetTab1State();
    this.setState({
      SelectedSites: sites,
      selectedSiteId: null,
      selectedTabValue: null,
      pageEntries: [],
      dateStartDate: new Date(),
      libraryEntries: [],
      selectedLibrary: null,
      isLibraryReportOpen: false,
      selectedFoundItem: null,
      isQueryingLibraries: false,
      isFilteringLibraries: false,
      chkShowLibaries: true,
      chkShowLists: true
    });
  }

  private resetTab1State(): void {
    this.setState({
      pageResults: [],
      selectedPage: null,
      isReportOpen: false,
      isProcessingBrokenLinks: false,
      expandedContentSections: new Set<string>(),
      showOnlyBrokenLinks: false,
      pageDetailsCache: new Map<string, PageStatusInfo>(),
      isLoadingPageDetails: false,
      pageDetailsLoaded: false,
      pageDetailsError: null,
      isPagePermissionsOpen: false,
      pagePermissions: [],
      isLoadingPagePermissions: false,
      pagePermissionsError: null,
      permissionGroupTree: [],
      openTreeNodeKeys: new Set<string>(),
      selectedTreeNodeKey: 'root',
      groupMemberCache: new Map<string, ResolvedGroupUser[]>(),
      isLoadingGroupMembers: false,
      groupMembersError: null,
      currentArtefact: null,
      permissionsSubjectTitle: '',
      permissionsSubjectUrl: null,
      isCheckingPrincipalAccess: false,
      principalAccessResult: null,
      principalAccessError: null
    });
  }

  private onDropdDownSelectionChanged = async (event: any, data: any): Promise<void> => {
    this.resetTab1State();
    const dataManager = new GraphDataManager(this.props.msGraphClientFactory, this.props.spHTTPClient);
    this.setState({ isFilteringLibraries: true });
    const pages = await dataManager.GetPages4Site(data.optionValue);
    this.setState({
      selectedTabValue: this.state.selectedTabValue === null ? "tab1" : this.state.selectedTabValue,
      pageEntries: pages,
      selectedSiteId: data.optionValue
    });
    try {
      const siteInfo: Site = this.state.SelectedSites.filter(x => x.id === data.optionValue)[0];
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

  private GetSelectedSite(): Site {
    return this.state.SelectedSites.filter(x => x.id === this.state.selectedSiteId)[0] as Site;
  }

  // Check-out is a document library feature (BaseType 1). Querying CheckoutUserId also errors
  // out on libraries that never had check-out enabled - ForceCheckout ("Require Check Out")
  // reliably indicates the feature is provisioned on the list, so it doubles as the gate here.
  private SupportsCheckout(listInfo: ListInformation): boolean {
    return listInfo.BaseType === 1 && !!listInfo.ForceCheckout;
  }

  private onTabSelect = (event: any, data: { value: TabValue }): void => {
    this.setState({ selectedTabValue: data.value });
  }

  private onPermissionsDialogTabSelect = (event: any, data: { value: TabValue }): void => {
    this.setState({ permissionsDialogTabValue: data.value });
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
