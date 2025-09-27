import { ListTemplateType } from "../../Core/ListTemplateTypes";

export interface ListInformation {
    AllowContentTypes: boolean;
    BaseTemplate: ListTemplateType;
    BaseType: number;
    ContentTypesEnabled: boolean;
    CrawlNonDefaultViews: boolean;
    Created: string;
    CurrentChangeToken: any; // SP.ChangeToken
    DefaultContentApprovalWorkflowId: string; // Guid
    DefaultItemOpenUseListSetting: boolean;
    Description?: string;
    Direction?: string;
    DisableCommenting: boolean;
    DisableGridEditing: boolean;
    DocumentTemplateUrl?: string;
    DraftVersionVisibility: number;
    EnableAttachments: boolean;
    EnableFolderCreation: boolean;
    EnableMinorVersions: boolean;
    EnableModeration: boolean;
    EnableRequestSignOff: boolean;
    EnableVersioning: boolean;
    EntityTypeName?: string;
    ExemptFromBlockDownloadOfNonViewableFiles: boolean;
    FileSavePostProcessingEnabled: boolean;
    ForceCheckout: boolean;
    HasExternalDataSource: boolean;
    Hidden: boolean;
    Id: string; // Guid
    ImagePath: any; // SP.ResourcePath
    ImageUrl?: string;
    DefaultSensitivityLabelForLibrary?: string;
    SensitivityLabelToEncryptOnDownloadForLibrary?: string | null;
    IrmEnabled: boolean;
    IrmExpire: boolean;
    IrmReject: boolean;
    IsApplicationList: boolean;
    IsCatalog: boolean;
    IsPrivate: boolean;
    ItemCount: number;
    LastItemDeletedDate: string;
    LastItemModifiedDate: string;
    LastItemUserModifiedDate: string;
    ListExperienceOptions: number;
    ListItemEntityTypeFullName?: string;
    MajorVersionLimit: number;
    MajorWithMinorVersionsLimit: number;
    MultipleDataList: boolean;
    NoCrawl: boolean;
    ParentWebPath: any; // SP.ResourcePath
    ParentWebUrl?: string;
    ParserDisabled: boolean;
    ServerTemplateCanCreateFolders: boolean;
    TemplateFeatureId: string; // Guid
    Title?: string;
  }