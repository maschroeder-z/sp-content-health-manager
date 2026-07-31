import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ContentHealthManagerWebPartStrings';
import ContentHealthManager from './components/ContentHealthManager';
import { IContentHealthManagerProps } from './components/IContentHealthManagerProps';
import { FluentProvider, FluentProviderProps, IdPrefixProvider, teamsLightTheme, teamsDarkTheme } from '@fluentui/react-components';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import './WebPartTitleOverrides.global.scss';

export interface IContentHealthManagerWebPartProps {
  description: string;
  title: string;
}

export enum AppMode {
  SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
}

export default class ContentHealthManagerWebPart extends BaseClientSideWebPart<IContentHealthManagerWebPartProps> {
  //private _appMode: AppMode = AppMode.SharePoint;
  //private _theme: Theme = webLightTheme;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IContentHealthManagerProps> = React.createElement(
    ContentHealthManager,
    {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      msGraphClientFactory: this.context.msGraphClientFactory as any,
      wpContext: this.context,
      spHTTPClient: this.context.spHttpClient as any      
    });


    const titleElement: React.ReactElement = React.createElement(WebPartTitle, {
      displayMode: this.displayMode,
      title: this.properties.title || strings.ContentHealthManagerTitle,
      updateProperty: (value: string) => {
        this.properties.title = value;
      }
    });

    const fluentElement: React.ReactElement<FluentProviderProps> = React.createElement(
      FluentProvider,
      {
        theme: this._isDarkTheme ? teamsDarkTheme : teamsLightTheme
      },
      titleElement,
      element
    );

    const temp: React.ReactElement = React.createElement(IdPrefixProvider,{value:"msz"},fluentElement);

    ReactDom.render(temp, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors,
      palette
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    if (palette) {
      this.domElement.style.setProperty('--themePrimary', palette.themePrimary || null);
      this.domElement.style.setProperty('--themeLighterAlt', palette.themeLighterAlt || null);
      this.domElement.style.setProperty('--themeLighter', palette.themeLighter || null);
      this.domElement.style.setProperty('--neutralLighter', palette.neutralLighter || null);
      this.domElement.style.setProperty('--neutralLight', palette.neutralLight || null);
      this.domElement.style.setProperty('--neutralSecondary', palette.neutralSecondary || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
