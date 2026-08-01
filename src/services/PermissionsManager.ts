import { ISPHttpClientOptions, MSGraphClientFactory, MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';
import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta";
import {
  PageStatusInfo,
  PermissionKind,
  PrincipalAccessReport,
  PrincipalReference,
  PrincipalType,
  ResolvedGroupUser,
  SharePointArtefact,
  SharePointArtefactType,
  SharePointGroupInfo,
  SharePointPermissionInfo,
  SharePointPrincipalPermission
} from '../models/REST/Permissions';

/**
 * Marks an error as an expected, already-explained condition (e.g. an Entra claim that turned
 * out to be a directory role rather than a group) so callers can log it quietly instead of as a
 * console.error - the message is already user-facing and actionable, not a bug to investigate.
 */
export class ExpectedPermissionResolutionError extends Error { }

export class PermissionsManager {
  private readonly spHttpClient: SPHttpClient;
  private readonly graphClientPromise: Promise<MSGraphClientV3>;

  constructor(msGraphClientFactory: MSGraphClientFactory, spHttpClient: SPHttpClient) {
    this.spHttpClient = spHttpClient;
    this.graphClientPromise = msGraphClientFactory.getClient('3');
  }

  public async get4ArtefactPermissions(artefact: SharePointArtefact): Promise<SharePointPrincipalPermission[]> {
    try {
      // A securable object's own `roleassignments` collection is empty while it inherits permissions
      // (HasUniqueRoleAssignments === false) - walk up to the nearest object with unique permissions
      // to get the effective set of principals.
      const base = await this.resolveNearestSecurableBaseUrl(artefact);
      const response = await this.spHttpClient.get(
        `${base}/roleassignments?$expand=Member,RoleDefinitionBindings`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const roleAssignments: any[] = this.unwrapCollection(data);

      return roleAssignments.map((roleAssignment: any) => {
        const member = roleAssignment.Member || {};
        const roles: any[] = roleAssignment.RoleDefinitionBindings?.results || [];
        const principalType: PrincipalType = member.PrincipalType;

        if (!member.LoginName && !member.Title) {
          console.warn('get4ArtefactPermissions: a role assignment\'s Member came back empty - the $expand=Member may have partially failed for this principal.', roleAssignment);
        }

        return {
          principalId: member.Id,
          principalType,
          isGroup: principalType !== PrincipalType.User,
          displayName: member.Title || member.LoginName || '(unresolved principal)',
          loginName: member.LoginName,
          email: member.Email || undefined,
          roles: roles.map((role: any) => role.Name)
        };
      });
    } catch (error) {
      console.error('Error retrieving artefact permissions:', error);
      throw error;
    }
  }

  public async checkPermission4User(user: MicrosoftGraphBeta.User, artefact: SharePointArtefact): Promise<SharePointPermissionInfo> {
    try {
      const loginHint = user.mail || user.userPrincipalName;
      if (!loginHint) {
        throw new Error('User must have a mail or userPrincipalName to resolve a SharePoint login name.');
      }

      const loginName = await this.resolveLoginName(artefact.webUrl, loginHint);
      return await this.getEffectivePermissions(loginName, artefact);
    } catch (error) {
      console.error('Error checking user permission:', error);
      throw error;
    }
  }

  /**
   * Resolves whether an arbitrary principal (found via a search UI) has effective access to an
   * artefact, and if so, its effective capability breakdown.
   */
  public async checkAccess4Principal(principal: PrincipalReference, artefact: SharePointArtefact): Promise<PrincipalAccessReport> {
    try {
      const loginName = /^\d+$/.test(principal.id)
        ? await this.resolveSharePointGroupLoginName(artefact.webUrl, Number(principal.id))
        : principal.id;

      const permissionInfo = await this.getEffectivePermissions(loginName, artefact);
      const hasAccess = permissionInfo.rawMask.high !== 0 || permissionInfo.rawMask.low !== 0;
      return { hasAccess, permissionInfo };
    } catch (error) {
      console.error('Error checking principal access:', error);
      throw error;
    }
  }

  private async getEffectivePermissions(loginName: string, artefact: SharePointArtefact): Promise<SharePointPermissionInfo> {
    const base = this.buildBaseUrl(artefact);
    const encodedLoginName = encodeURIComponent(`'${loginName.replace(/'/g, "''")}'`);
    const response = await this.spHttpClient.get(
      `${base}/getusereffectivepermissions(@v)?@v=${encodedLoginName}`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=verbose' } }
    );

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    const entity = this.unwrapEntity(data);
    const mask = {
      high: Number(entity?.High || 0),
      low: Number(entity?.Low || 0)
    };

    return {
      loginName,
      canView: this.hasPermissionKind(mask, PermissionKind.ViewListItems),
      canContribute: this.hasPermissionKind(mask, PermissionKind.AddListItems) && this.hasPermissionKind(mask, PermissionKind.EditListItems),
      canEdit: this.hasPermissionKind(mask, PermissionKind.EditListItems),
      canDelete: this.hasPermissionKind(mask, PermissionKind.DeleteListItems),
      canManageLists: this.hasPermissionKind(mask, PermissionKind.ManageLists),
      canManagePermissions: this.hasPermissionKind(mask, PermissionKind.ManagePermissions),
      hasFullControl: this.hasPermissionKind(mask, PermissionKind.FullMask),
      rawMask: mask
    };
  }

  private async resolveSharePointGroupLoginName(webUrl: string, groupId: number): Promise<string> {
    const response = await this.spHttpClient.get(
      `${this.normalizeWebUrl(webUrl)}/_api/web/sitegroups/getbyid(${groupId})?$select=LoginName`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=verbose' } }
    );

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    const loginName = this.unwrapEntity(data)?.LoginName;
    if (!loginName) {
      throw new Error(`Could not resolve LoginName for SharePoint group ${groupId}.`);
    }
    return loginName;
  }

  /**
   * Resolves the SharePointArtefact (list + item) that owns the file/page at the given URL.
   * Needed because identifiers exposed by other APIs (e.g. Microsoft Graph's sitePage id) do not
   * necessarily match the underlying SharePoint list item id.
   */
  public async resolveArtefactFromFileUrl(webUrl: string, fileUrl: string): Promise<SharePointArtefact> {
    try {
      const normalizedWebUrl = this.normalizeWebUrl(webUrl);
      const serverRelativeUrl = this.toServerRelativeUrl(fileUrl);

      const encodedServerRelativeUrl = serverRelativeUrl
        .replace(/'/g, "''")
        .replace(/\(/g, '%28')
        .replace(/\)/g, '%29');

      const response = await this.spHttpClient.get(
        `${normalizedWebUrl}/_api/web/GetFileByServerRelativeUrl('${encodedServerRelativeUrl}')/ListItemAllFields?$select=Id,ParentList/Id&$expand=ParentList`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const entity = this.unwrapEntity(data);
      const listId = entity?.ParentList?.Id;
      const itemId = Number(entity?.Id);
      if (!listId || Number.isNaN(itemId)) {
        throw new Error(`Could not resolve list item for file URL: ${fileUrl}`);
      }

      return {
        type: SharePointArtefactType.ListItem,
        webUrl: normalizedWebUrl,
        listId,
        itemId
      };
    } catch (error) {
      console.error('Error resolving artefact from file URL:', error);
      throw error;
    }
  }

  /**
   * Retrieves the extra per-page status shown by the Pages overview list's "Load details" action:
   * whether the page has an unpublished draft, whether it has unique (non-inherited) permissions,
   * and who (if anyone) has it checked out. Fetched in a single REST call by extending the same
   * GetFileByServerRelativeUrl/ListItemAllFields query shape used by resolveArtefactFromFileUrl.
   */
  public async getPageStatus(webUrl: string, fileUrl: string): Promise<PageStatusInfo> {
    try {
      const normalizedWebUrl = this.normalizeWebUrl(webUrl);
      const serverRelativeUrl = this.toServerRelativeUrl(fileUrl);

      const encodedServerRelativeUrl = serverRelativeUrl
        .replace(/'/g, "''")
        .replace(/\(/g, '%28')
        .replace(/\)/g, '%29');

      const response = await this.spHttpClient.get(
        `${normalizedWebUrl}/_api/web/GetFileByServerRelativeUrl('${encodedServerRelativeUrl}')/ListItemAllFields` +
        `?$select=HasUniqueRoleAssignments,File/Level,File/CheckOutType,File/CheckedOutByUser/Title` +
        `&$expand=File,File/CheckedOutByUser`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const entity = this.unwrapEntity(data);
      const file = entity?.File;
      // SP.CheckOutType: 0 = Online, 1 = Offline, 2 = None (not checked out)
      const checkedOutBy = (file && file.CheckOutType !== 2 && file.CheckedOutByUser)
        ? file.CheckedOutByUser.Title
        : null;

      return {
        needsApproval: file?.Level === 1, //'Draft',
        hasUniquePermission: !!entity?.HasUniqueRoleAssignments,
        checkedOutBy
      };
    } catch (error) {
      console.error('Error retrieving page status:', error);
      throw error;
    }
  }

  /**
   * Resolves the site's pages library (the "Site Pages" list, BaseTemplate 119 - WebPageLibrary) as a
   * List-type artefact, for checking permissions on the library itself rather than a single page.
   * Filtering by BaseTemplate rather than title keeps this locale-independent.
   */
  public async resolvePagesLibraryArtefact(webUrl: string): Promise<SharePointArtefact> {
    try {
      const normalizedWebUrl = this.normalizeWebUrl(webUrl);
      const response = await this.spHttpClient.get(
        `${normalizedWebUrl}/_api/web/lists?$filter=BaseTemplate eq 119&$select=Id&$top=1`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const lists = this.unwrapCollection(data);
      const listId = lists?.[0]?.Id;
      if (!listId) {
        throw new Error('Could not resolve the pages library for this site.');
      }

      return {
        type: SharePointArtefactType.List,
        webUrl: normalizedWebUrl,
        listId
      };
    } catch (error) {
      console.error('Error resolving pages library artefact:', error);
      throw error;
    }
  }

  public async hasUniquePermission(artefact: SharePointArtefact): Promise<boolean> {
    try {
      return await this.getHasUniqueRoleAssignments(this.buildBaseUrl(artefact));
    } catch (error) {
      console.error('Error checking unique permissions:', error);
      throw error;
    }
  }

  /**
   * Resolves all users belonging to a group principal, whether it is a native SharePoint group
   * (resolved via SharePoint REST) or an Entra ID group (resolved via Microsoft Graph, since Entra
   * groups are not queryable through SharePoint REST).
   */
  public async resolveUser4Group(groupInfo: SharePointGroupInfo): Promise<ResolvedGroupUser[]> {
    try {
      if (groupInfo.principalType === PrincipalType.SharePointGroup) {
        if (!groupInfo.principalId) {
          throw new Error('groupInfo.principalId is required to resolve a SharePoint group.');
        }
        return await this.resolveSharePointGroupUsers(groupInfo.webUrl, groupInfo.principalId);
      }

      const entraGroupId = this.tryExtractEntraGroupId(groupInfo.loginName);
      if (entraGroupId) {
        return await this.resolveEntraGroupUsers(entraGroupId);
      }

      throw new Error(`Unsupported group principal (not a SharePoint group or Entra-backed security group): ${groupInfo.loginName || groupInfo.displayName}`);
    } catch (error) {
      if (error instanceof ExpectedPermissionResolutionError) {
        console.warn('Could not resolve group users:', error.message);
      } else {
        console.error('Error resolving group users:', error);
      }
      throw error;
    }
  }

  /**
   * Resolves the direct child groups of a group principal (one level, not transitive) — unlike
   * resolveUser4Group, which resolves effective/transitive users and does not surface nested-group
   * structure. Used to lazily populate a group hierarchy (e.g. a tree UI) one expansion at a time.
   */
  public async resolveNestedGroups(groupInfo: SharePointGroupInfo): Promise<SharePointGroupInfo[]> {
    try {
      if (groupInfo.principalType === PrincipalType.SharePointGroup) {
        if (!groupInfo.principalId) {
          throw new Error('groupInfo.principalId is required to resolve a SharePoint group.');
        }
        return await this.resolveSharePointGroupNestedGroups(groupInfo.webUrl, groupInfo.principalId);
      }

      const entraGroupId = this.tryExtractEntraGroupId(groupInfo.loginName);
      if (entraGroupId) {
        return await this.resolveEntraGroupNestedGroups(groupInfo.webUrl, entraGroupId);
      }

      throw new Error(`Unsupported group principal (not a SharePoint group or Entra-backed security group): ${groupInfo.loginName || groupInfo.displayName}`);
    } catch (error) {
      if (error instanceof ExpectedPermissionResolutionError) {
        console.warn('Could not resolve nested groups:', error.message);
      } else {
        console.error('Error resolving nested groups:', error);
      }
      throw error;
    }
  }

  /**
   * Resolves the members of a built-in Entra directory role (e.g. Global Administrator), given its
   * universal role template id (see models/REST/Permissions.ts SHAREPOINT_RELEVANT_ENTRA_ROLES).
   * Unlike security groups, directory roles are addressed via /directoryRoles, not /groups - a role
   * that has never been assigned to anyone in this tenant may not be "activated" yet and this call
   * 404s for it (activating one requires the write-scoped RoleManagement.ReadWrite.Directory
   * permission, out of scope for a read-only permissions audit tool). That 404 is treated as "nobody
   * holds this role" rather than an error and resolves to an empty list. Requires the
   * RoleManagement.Read.Directory Graph permission, requested in config/package-solution.json and
   * subject to tenant admin approval - until approved this throws a 403 GraphError.
   */
  public async resolveDirectoryRoleUsers(roleTemplateId: string): Promise<ResolvedGroupUser[]> {
    try {
      const client = await this.graphClientPromise;

      const role = await client
        .api(`/directoryRoles(roleTemplateId='${encodeURIComponent(roleTemplateId)}')`)
        .version('v1.0')
        .select('id')
        .get();

      return await this.fetchDirectoryRoleMembersByRoleId(client, role.id);
    } catch (error: any) {
      if (error?.statusCode === 404) {
        return [];
      }
      console.error('Error resolving directory role members:', error);
      throw error;
    }
  }

  /**
   * Fetches the members of an already-activated directory role, given its own (tenant-specific)
   * object id - not a roleTemplateId. Shared by resolveDirectoryRoleUsers (which resolves a
   * universal roleTemplateId to this id first) and resolveEntraGroupUsers's directory-role
   * fallback (which already has what's presumably this id, extracted directly from a SharePoint
   * claim, so it skips the roleTemplateId resolution step entirely).
   */
  private async fetchDirectoryRoleMembersByRoleId(client: MSGraphClientV3, roleId: string): Promise<ResolvedGroupUser[]> {
    const response = await client
      .api(`/directoryRoles/${encodeURIComponent(roleId)}/members`)
      .version('v1.0')
      .select(['id', 'displayName', 'mail', 'userPrincipalName'].join(','))
      .get();

    const members: any[] = response?.value || [];
    return members.map((member: any) => ({
      id: member.id,
      displayName: member.displayName,
      email: member.mail || member.userPrincipalName || undefined
    }));
  }

  private async resolveSharePointGroupUsers(webUrl: string, groupId: number): Promise<ResolvedGroupUser[]> {
    const members = await this.fetchSharePointGroupMembers(webUrl, groupId);

    return members
      .filter((member: any) => !this.tryExtractEntraGroupId(member.LoginName))
      .map((member: any) => ({
        id: String(member.Id),
        displayName: member.Title || member.LoginName || '(unresolved principal)',
        email: member.Email || undefined,
        loginName: member.LoginName
      }));
  }

  /**
   * A SharePoint group cannot contain another SharePoint group, but an Entra group can be added as a
   * member of one — it then shows up in the group's `users` collection as a `c:0t.c|tenant|<guid>`
   * login name. SharePoint's legacy `sitegroups(id)/users` endpoint reports this shadow entity's
   * PrincipalType as a plain User, so the login-name claims pattern (not PrincipalType) is the only
   * reliable way to recognize it here.
   */
  private async resolveSharePointGroupNestedGroups(webUrl: string, groupId: number): Promise<SharePointGroupInfo[]> {
    const members = await this.fetchSharePointGroupMembers(webUrl, groupId);

    return members
      .filter((member: any) => !!this.tryExtractEntraGroupId(member.LoginName))
      .map((member: any) => ({
        webUrl,
        principalId: member.Id,
        principalType: PrincipalType.SecurityGroup,
        loginName: member.LoginName,
        displayName: member.Title
      }));
  }

  /**
   * SharePoint REST collections default to a ~100-item page unless $top is specified, and this
   * endpoint doesn't automatically follow up - without pagination, groups larger than that would
   * silently lose members past the first page. $top=5000 matches SharePoint's own hard collection
   * cap, and d.__next (odata=verbose) is followed in case a tenant enforces a smaller page size.
   */
  private async fetchSharePointGroupMembers(webUrl: string, groupId: number): Promise<any[]> {
    const members: any[] = [];
    let nextUrl: string | undefined =
      `${this.normalizeWebUrl(webUrl)}/_api/web/sitegroups/getbyid(${groupId})/users?$top=5000`;

    while (nextUrl) {
      const response = await this.spHttpClient.get(
        nextUrl,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      members.push(...this.unwrapCollection(data));
      nextUrl = data?.d?.__next || undefined;
    }

    const missingDetails = members.filter((member: any) => !member.LoginName || !member.Title);
    if (missingDetails.length > 0) {
      console.warn(`fetchSharePointGroupMembers: ${missingDetails.length} member(s) of group ${groupId} came back with a missing LoginName/Title - the REST expand may have partially failed for these principals.`, missingDetails);
    }

    return members;
  }

  private async resolveEntraGroupUsers(groupId: string): Promise<ResolvedGroupUser[]> {
    const client = await this.graphClientPromise;

    try {
      const response = await client
        .api(`/groups/${encodeURIComponent(groupId)}/transitiveMembers/microsoft.graph.user`)
        .version('v1.0')
        .select(['id', 'displayName', 'mail', 'userPrincipalName'].join(','))
        .get();

      const members: any[] = response?.value || [];
      return members.map((member: any) => ({
        id: member.id,
        displayName: member.displayName,
        email: member.mail || member.userPrincipalName || undefined
      }));
    } catch (error: any) {
      if (error?.statusCode === 404) {
        try {
          return await this.fetchDirectoryRoleMembersByRoleId(client, groupId);
        } catch (roleError: any) {
          if (roleError?.statusCode === 403) {
            throw new ExpectedPermissionResolutionError(
              `This Entra ID object (${groupId}) appears to be a built-in directory role, but your tenant admin has not approved this app's RoleManagement.Read.Directory permission request yet - approve it in the SharePoint Admin Center's API access page to list its members.`
            );
          }
          // Not resolvable as a directory role either - fall through to the generic message below.
        }
      }
      throw this.toExpectedGraphGroupError(error, groupId);
    }
  }

  /**
   * SharePoint represents both real Entra security groups and built-in Entra directory roles
   * (Global Administrator, SharePoint Administrator, etc.) with the same claim format
   * (`c:0t.c|tenant|<guid>`), so a claim that's actually a role id (or a since-deleted group)
   * 404s here - Graph's /groups endpoint doesn't recognize either as a group. Surface that as a
   * clear, expected condition instead of letting Graph's raw error propagate as a console-error
   * storm; a 403 means the same claim resolved but this app isn't authorized to read it (e.g. a
   * role membership read that needs a permission this app hasn't been granted yet).
   */
  private toExpectedGraphGroupError(error: any, groupId: string): Error {
    const statusCode = error?.statusCode;
    if (statusCode === 404) {
      return new ExpectedPermissionResolutionError(
        `This Entra ID object (${groupId}) couldn't be found as a security group - it may be a built-in directory role (like Global Administrator), which SharePoint represents using the same claim format, or a security group that has since been deleted. Role membership can't be listed here without the RoleManagement.Read.Directory permission.`
      );
    }
    if (statusCode === 403) {
      return new ExpectedPermissionResolutionError(
        `Access denied reading Entra ID object ${groupId} - your tenant admin has not approved this app's pending Graph API permission request yet. Approve it in the SharePoint Admin Center's API access page to continue.`
      );
    }
    return error instanceof Error ? error : new Error(String(error));
  }

  /** Direct (non-transitive) nested groups of an Entra group, one level down. */
  private async resolveEntraGroupNestedGroups(webUrl: string, groupId: string): Promise<SharePointGroupInfo[]> {
    const client = await this.graphClientPromise;

    try {
      return await this.fetchEntraGroupNestedGroups(client, webUrl, groupId);
    } catch (error: any) {
      if (error?.statusCode === 404) {
        try {
          return await this.fetchDirectoryRoleNestedGroups(client, webUrl, groupId);
        } catch (roleError: any) {
          if (roleError?.statusCode === 403) {
            throw new ExpectedPermissionResolutionError(
              `This Entra ID object (${groupId}) appears to be a built-in directory role, but your tenant admin has not approved this app's RoleManagement.Read.Directory permission request yet - approve it in the SharePoint Admin Center's API access page to list its nested groups.`
            );
          }
          // Not resolvable as a directory role either - fall through to the generic message below.
        }
      }
      throw this.toExpectedGraphGroupError(error, groupId);
    }
  }

  private async fetchEntraGroupNestedGroups(client: MSGraphClientV3, webUrl: string, groupId: string): Promise<SharePointGroupInfo[]> {
    const response = await client
      .api(`/groups/${encodeURIComponent(groupId)}/members/microsoft.graph.group`)
      .version('v1.0')
      .select(['id', 'displayName'].join(','))
      .get();

    const members: any[] = response?.value || [];
    return members.map((member: any) => ({
      webUrl,
      principalType: PrincipalType.SecurityGroup,
      loginName: `c:0t.c|tenant|${member.id}`,
      displayName: member.displayName
    }));
  }

  /**
   * Directory roles are usually assigned to users, but Entra RBAC also allows assigning a role to a
   * group (its members then inherit the role) - so, mirroring fetchEntraGroupNestedGroups, this
   * asks Graph for the group-typed members of an already-activated role, given its own object id.
   */
  private async fetchDirectoryRoleNestedGroups(client: MSGraphClientV3, webUrl: string, roleId: string): Promise<SharePointGroupInfo[]> {
    const response = await client
      .api(`/directoryRoles/${encodeURIComponent(roleId)}/members/microsoft.graph.group`)
      .version('v1.0')
      .select(['id', 'displayName'].join(','))
      .get();

    const members: any[] = response?.value || [];
    return members.map((member: any) => ({
      webUrl,
      principalType: PrincipalType.SecurityGroup,
      loginName: `c:0t.c|tenant|${member.id}`,
      displayName: member.displayName
    }));
  }

  /** Extracts the AAD group id from a claims-encoded Entra group login name, e.g. `c:0t.c|tenant|<guid>`. */
  private tryExtractEntraGroupId(loginName?: string): string | undefined {
    if (!loginName) {
      return undefined;
    }
    const match = /^c:0t\.c\|tenant\|([0-9a-fA-F-]{36})$/.exec(loginName);
    return match ? match[1] : undefined;
  }

  /**
   * Walks up Web -> List -> ListItem until it finds the object that actually owns the permissions
   * (HasUniqueRoleAssignments === true), since an inheriting object's own `roleassignments` collection
   * is empty rather than mirroring its parent's.
   */
  private async resolveNearestSecurableBaseUrl(artefact: SharePointArtefact): Promise<string> {
    let current: SharePointArtefact = artefact;
    for (; ;) {
      const base = this.buildBaseUrl(current);
      if (current.type === SharePointArtefactType.Web || await this.getHasUniqueRoleAssignments(base)) {
        return base;
      }
      current = current.type === SharePointArtefactType.ListItem
        ? { type: SharePointArtefactType.List, webUrl: current.webUrl, listId: current.listId }
        : { type: SharePointArtefactType.Web, webUrl: current.webUrl };
    }
  }

  private async getHasUniqueRoleAssignments(base: string): Promise<boolean> {
    const response = await this.spHttpClient.get(
      `${base}?$select=HasUniqueRoleAssignments`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=verbose' } }
    );

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return !!this.unwrapEntity(data)?.HasUniqueRoleAssignments;
  }

  private async resolveLoginName(webUrl: string, loginHint: string): Promise<string> {
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      body: JSON.stringify({ logonName: loginHint })
    };

    try {
      const response = await this.spHttpClient.post(
        `${this.normalizeWebUrl(webUrl)}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        options
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      return this.unwrapEntity(data)?.LoginName;
    } catch (error) {
      throw new Error(`User could not be resolved on this web (${loginHint}): ${error}`);
    }
  }

  /** Normalizes an odata=verbose entity response (`{ d: {...} }`) or a minimal/nometadata one (the object itself). */
  private unwrapEntity(data: any): any {
    return data?.d ?? data;
  }

  /** Normalizes an odata=verbose collection response (`{ d: { results: [...] } }`) or a minimal/nometadata one (`{ value: [...] }`). */
  private unwrapCollection(data: any): any[] {
    if (data?.d?.results) {
      return data.d.results;
    }
    if (Array.isArray(data?.value)) {
      return data.value;
    }
    return [];
  }

  private normalizeWebUrl(webUrl: string): string {
    return webUrl.replace(/\/+$/, '');
  }

  private toServerRelativeUrl(url: string): string {
    if (!/^https?:\/\//i.test(url)) {
      return url;
    }
    const withoutProtocol = url.replace(/^https?:\/\//i, '');
    const firstSlash = withoutProtocol.indexOf('/');
    return firstSlash === -1 ? '/' : decodeURIComponent(withoutProtocol.substring(firstSlash));
  }

  private buildBaseUrl(artefact: SharePointArtefact): string {
    const webUrl = this.normalizeWebUrl(artefact.webUrl);

    switch (artefact.type) {
      case SharePointArtefactType.Web:
        return `${webUrl}/_api/web`;
      case SharePointArtefactType.List:
        if (!artefact.listId) {
          throw new Error('artefact.listId is required for SharePointArtefactType.List');
        }
        return `${webUrl}/_api/web/lists('${artefact.listId}')`;
      case SharePointArtefactType.ListItem:
        if (!artefact.listId || artefact.itemId === undefined) {
          throw new Error('artefact.listId and artefact.itemId are required for SharePointArtefactType.ListItem');
        }
        return `${webUrl}/_api/web/lists('${artefact.listId}')/items(${artefact.itemId})`;
      default:
        throw new Error(`Unsupported SharePointArtefactType: ${artefact.type}`);
    }
  }

  private hasPermissionKind(mask: { high: number; low: number }, kind: PermissionKind): boolean {
    if (kind === PermissionKind.FullMask) {
      return (mask.high & 0x7FFFFFFF) === 0x7FFFFFFF && mask.low === 0xFFFFFFFF;
    }

    const bit = kind - 1;
    if (bit < 32) {
      return (mask.low & (1 << bit)) !== 0;
    }
    return (mask.high & (1 << (bit - 32))) !== 0;
  }
}

export default PermissionsManager;
