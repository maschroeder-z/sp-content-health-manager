export enum SharePointArtefactType {
  Web = 'Web',
  List = 'List',
  ListItem = 'ListItem'
}

/** Generic reference to any SharePoint securable object (web, list/library, or list item/document/page/image). */
export interface SharePointArtefact {
  type: SharePointArtefactType;
  webUrl: string;     // absolute URL of the owning web, no trailing slash
  listId?: string;    // list/library GUID; required when type is List or ListItem
  itemId?: number;    // list item id; required when type is ListItem
}

export enum PrincipalType {
  None = 0,
  User = 1,
  DistributionList = 2,
  SecurityGroup = 4,
  SharePointGroup = 8,
  All = 15
}

/** One entry of the result of PermissionsManager.get4ArtefactPermissions. */
export interface SharePointPrincipalPermission {
  principalId: number;        // SP.User/SP.Group numeric Id (site-scoped, not an AAD id)
  principalType: PrincipalType;
  isGroup: boolean;
  displayName: string;
  loginName: string;          // claims-encoded for users, e.g. i:0#.f|membership|user@tenant.com
  email?: string;
  roles: string[];            // RoleDefinitionBindings[].Name, e.g. ["Full Control", "Read"]
}

// Curated subset of SP.PermissionKind (real CSOM numeric values, 1-based bit positions).
export enum PermissionKind {
  ViewListItems = 1,
  AddListItems = 2,
  EditListItems = 3,
  DeleteListItems = 4,
  ManageLists = 12,
  ViewPages = 18,
  ManagePermissions = 26,
  ManageWeb = 31,
  FullMask = 65
}

/** Result of PermissionsManager.checkPermission4User. */
export interface SharePointPermissionInfo {
  loginName: string;             // resolved claims login name used for the check
  canView: boolean;
  canContribute: boolean;        // AddListItems && EditListItems
  canEdit: boolean;
  canDelete: boolean;
  canManageLists: boolean;
  canManagePermissions: boolean;
  hasFullControl: boolean;
  rawMask: { high: number; low: number };
}
