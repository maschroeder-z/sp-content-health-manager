import { MSGraphClientFactory } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IContentHealthManagerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  wpContext: WebPartContext;
  msGraphClientFactory: MSGraphClientFactory;
}
