import { MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IContentHealthManagerProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  wpContext: WebPartContext;
  msGraphClientFactory: MSGraphClientFactory;
  spHTTPClient: SPHttpClient;
}
