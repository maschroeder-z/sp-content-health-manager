import { ISPHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';
import * as MicrosoftGraphBeta from "@microsoft/microsoft-graph-types-beta";
import {
  PermissionKind,
  PrincipalType,
  SharePointArtefact,
  SharePointArtefactType,
  SharePointPermissionInfo,
  SharePointPrincipalPermission
} from '../models/REST/Permissions';

export class PermissionsManager {
  private readonly spHttpClient: SPHttpClient;

  constructor(spHttpClient: SPHttpClient) {
    this.spHttpClient = spHttpClient;
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

        return {
          principalId: member.Id,
          principalType,
          isGroup: principalType !== PrincipalType.User,
          displayName: member.Title,
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
    } catch (error) {
      console.error('Error checking user permission:', error);
      throw error;
    }
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

  public async hasUniquePermission(artefact: SharePointArtefact): Promise<boolean> {
    try {
      return await this.getHasUniqueRoleAssignments(this.buildBaseUrl(artefact));
    } catch (error) {
      console.error('Error checking unique permissions:', error);
      throw error;
    }
  }

  /**
   * Walks up Web -> List -> ListItem until it finds the object that actually owns the permissions
   * (HasUniqueRoleAssignments === true), since an inheriting object's own `roleassignments` collection
   * is empty rather than mirroring its parent's.
   */
  private async resolveNearestSecurableBaseUrl(artefact: SharePointArtefact): Promise<string> {
    let current: SharePointArtefact = artefact;
    for (;;) {
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
