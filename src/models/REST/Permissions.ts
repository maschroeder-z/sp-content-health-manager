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

/** Input to PermissionsManager.resolveUser4Group. Compatible with a SharePointPrincipalPermission
 *  plus the webUrl it was scoped to (SP groups are web-scoped, so the REST call needs it). */
export interface SharePointGroupInfo {
  webUrl: string;              // absolute URL of the web the group is scoped to
  principalId?: number;        // SP.Group Id (site-scoped); required when principalType is SharePointGroup
  principalType: PrincipalType;
  loginName?: string;          // claims login name; required to detect/resolve an Entra group, e.g. c:0t.c|tenant|<aadGroupId>
  displayName?: string;
}

/** One user resolved from a SharePoint or Entra group. */
export interface ResolvedGroupUser {
  id: string;                  // AAD object id (Entra-resolved) or SP.User Id as string (SharePoint-group-resolved)
  displayName: string;
  email?: string;
  loginName?: string;          // SharePoint claims login name; present for SharePoint-group-resolved users
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

/** A principal picked from a search UI (SharePoint REST people/group picker). */
export interface PrincipalReference {
  id: string;           // claims login name (user or Entra/security group) OR a numeric SharePoint group id
  displayName: string;
}

/** Result of PermissionsManager.checkAccess4Principal. */
export interface PrincipalAccessReport {
  hasAccess: boolean;
  permissionInfo: SharePointPermissionInfo;
}

/** Extra per-page status fetched on demand for the Pages overview list's "Load details" action. */
export interface PageStatusInfo {
  needsApproval: boolean;        // SP.File Level === 'Draft': saved changes not yet published
  hasUniquePermission: boolean;
  checkedOutBy: string | null;   // display name of the user with the page checked out, or null if not checked out
}
