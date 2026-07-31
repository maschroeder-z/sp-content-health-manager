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

/** One entry in the built-in Entra directory role picker (PermissionsManager.resolveDirectoryRoleUsers). */
export interface DirectoryRoleOption {
  displayName: string;
  roleTemplateId: string;  // universal Entra role template GUID, same across every tenant
}

/**
 * Built-in Entra directory roles relevant to a SharePoint permissions audit, with their universal
 * role template GUIDs (identical in every tenant). Verified against Microsoft's built-in roles
 * reference: https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference
 * Querying a role's members requires the RoleManagement.Read.Directory Graph permission, requested
 * in config/package-solution.json and subject to tenant admin approval.
 */
export const SHAREPOINT_RELEVANT_ENTRA_ROLES: DirectoryRoleOption[] = [
  { displayName: 'Global Administrator', roleTemplateId: '62e90394-69f5-4237-9190-012177145e10' },
  { displayName: 'SharePoint Administrator', roleTemplateId: 'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' },
  { displayName: 'Global Reader', roleTemplateId: 'f2ef992c-3afb-46b9-b7cf-a126ee74c451' },
  { displayName: 'User Administrator', roleTemplateId: 'fe930be7-5e62-47db-91af-98c3a49a38b1' },
  { displayName: 'Security Administrator', roleTemplateId: '194ae4cb-b126-40b2-bd5b-6091b380977d' },
  { displayName: 'Teams Administrator', roleTemplateId: '69091246-20e8-4a56-aa4d-066075b2a7a8' },
  { displayName: 'Exchange Administrator', roleTemplateId: '29232cdf-9323-42fd-ade2-1d097af3e4de' },
  { displayName: 'Compliance Administrator', roleTemplateId: '17315797-102d-40b4-93e0-432062caca18' }
];
