export interface SPView {
    __metadata: Metadata
    ViewFields: ViewFields
    Aggregations: any
    AggregationsStatus: any
    AssociatedContentTypeId: any
    BaseViewId: string
    CalendarViewStyles: any
    ColumnWidth: any
    ContentTypeId: ContentTypeId
    CustomFormatter: any
    DefaultView: boolean
    DefaultViewForContentType: boolean
    EditorModified: boolean
    Formats: any
    GridInitInfo: GridInitInfo
    GridLayout: any
    Hidden: boolean
    HtmlSchemaXml: string
    Id: string
    ImageUrl: string
    IncludeRootFolder: boolean
    ViewJoins: any
    JSLink: any
    ListViewXml: string
    Method: any
    MobileDefaultView: boolean
    MobileView: boolean
    ModerationType: any
    NewDocumentTemplates: any
    OrderedView: boolean
    Paged: boolean
    PersonalView: boolean
    ViewProjectedFields: any
    ViewQuery: string
    ReadOnlyView: boolean
    RequiresClientIntegration: boolean
    RowLimit: number
    Scope: number
    ServerRelativePath: ServerRelativePath
    ServerRelativeUrl: string
    StyleId: any
    TabularView: boolean
    Threaded: boolean
    Title: string
    Toolbar: any
    ToolbarTemplateName: any
    ViewType: string
    ViewData: any
    ViewType2: any
    VisualizationInfo: any
  }
  
  export interface Metadata {
    id: string
    uri: string
    type: string
  }
  
  export interface ViewFields {
    __deferred: Deferred
  }
  
  export interface Deferred {
    uri: string
  }
  
  export interface ContentTypeId {
    __metadata: Metadata2
    StringValue: string
  }
  
  export interface Metadata2 {
    type: string
  }
  
  export interface GridInitInfo {
    __metadata: Metadata3
    ControllerId: any
    GridSerializer: any
    JsInitObj: any
  }
  
  export interface Metadata3 {
    type: string
  }
  
  export interface ServerRelativePath {
    __metadata: Metadata4
    DecodedUrl: string
  }
  
  export interface Metadata4 {
    type: string
  }
  