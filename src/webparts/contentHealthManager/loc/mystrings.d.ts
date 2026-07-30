declare interface IContentHealthManagerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  // Content Health Manager
  ContentHealthManagerTitle: string;
  ContentHealthManagerDescription: string;
  HowToUseHeading: string;
  FirstSelectSites: string;
  SecondSelectSingleSite: string;
  StartQueryToFind: string;
  BrokenLinksInPages: string;
  OldContentForDate: string;
  CheckedOutContentItems: string;
  PagesWaitingForApproval: string;
  FourthCheckPermissions: string;
  // Site Selection
  SelectFirstAllSites: string;
  SelectSitesLabel: string;
  SelectAllSitesPlaceholder: string;
  FilterSitesPlaceholder: string;
  ToContinueSelectSite: string;
  ChooseSiteLabel: string;
  SelectSitePlaceholder: string;
  ResultsForSite: string;
  // Tabs
  BrokenLinksAnalysisTab: string;
  LibraryAnalysisTab: string;
  // Library Analysis
  SelectDateLabel: string;
  SelectQueryDatePlaceholder: string;
  SelectDateHint: string;
  QueryAllLibraries: string;
  QueryLibrary: string;
  CheckedOutItems: string;
  OpenDetails: string;
  LibrariesCheckbox: string;
  ListsCheckbox: string;
  // Broken Links Analysis
  FindBrokenLinks: string;
  ProcessPage: string;
  // Button tooltips
  TooltipQueryAllLibraries: string;
  TooltipQueryLibrary: string;
  TooltipCheckedOutItems: string;
  TooltipOpenLibraryDetails: string;
  TooltipFindBrokenLinks: string;
  TooltipProcessPage: string;
  TooltipOpenPageDetails: string;
  TooltipShowPermissions: string;
  TooltipCloseDialog: string;
  TooltipToggleContent: string;
  TooltipLoadPageDetails: string;
  TooltipShowSelectedLibraryPermissions: string;
  // Page Report Dialog
  PageReportTitle: string;
  TitleLabel: string;
  UrlLabel: string;
  TotalLinksLabel: string;
  BrokenLinksLabel: string;
  AllLinksLabel: string;
  ShowOnlyBrokenLinks: string;
  NoTitle: string;
  ShowContent: string;
  NoBrokenLinksFound: string;
  NoLinksFound: string;
  NoLinkAnalysisAvailable: string;
  NoItemSelected: string;
  CloseButton: string;
  // Library Report Dialog
  LibraryReportTitle: string;
  TemplateLabel: string;
  DescriptionLabel: string;
  ItemCountLabel: string;
  CreatedLabel: string;
  LastModifiedLabel: string;
  LastUserModifiedLabel: string;
  LastDeletedLabel: string;
  EnableVersioningLabel: string;
  EnableAttachmentsLabel: string;
  EnableFolderCreationLabel: string;
  OverviewListEntries: string;
  TotalItemsFound: string;
  TotalCheckedOutIemsFound: string;
  ShowPermissions: string;
  QueryLibraryForResults: string;
  NoLibrarySelected: string;
  Yes: string;
  No: string;
  NA: string;
  // Page Permissions Dialog
  PermissionsButtonLabel: string;
  PagePermissionsTitle: string;
  PagesLibraryLabel: string;
  TooltipShowLibraryPermissions: string;
  PrincipalNameLabel: string;
  PrincipalTypeLabel: string;
  LoginNameLabel: string;
  RolesLabel: string;
  GroupLabel: string;
  UserLabel: string;
  NoPermissionsFound: string;
  NoNestedGroups: string;
  EmailLabel: string;
  SearchUserOrGroupPlaceholder: string;
  HasAccessLabel: string;
  NoAccessLabel: string;
  FullControlLabel: string;
  DesignLabel: string;
  EditLabel: string;
  ContributeLabel: string;
  ReadLabel: string;
  NoAccessLevelLabel: string;
  // View Fields
  FoundLabel: string;
  StartQueryForResults: string;
  CreatedAtLabel: string;
  LastChangeLabel: string;
  UserChangedLabel: string;
  LastDeletionLabel: string;
  FoundLinksCount: string;
  BrokenLinksCount: string;
  // Page Details (on-demand columns)
  LoadPageDetailsButtonLabel: string;
  NeedsApprovalLabel: string;
  HasUniquePermissionLabel: string;
  CheckedOutLabel: string;
  NotCheckedOut: string;
  CheckoutNotSupported: string;
  LibraryTypeLabel: string;
  ListTypeLabel: string;
}

declare module 'ContentHealthManagerWebPartStrings' {
  const strings: IContentHealthManagerWebPartStrings;
  export = strings;
}
