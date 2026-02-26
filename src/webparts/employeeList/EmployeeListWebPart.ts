import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  type IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { SPHttpClient } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EmployeeListWebPartStrings';
import EmployeeList from './components/EmployeeList';
import { IEmployeeListProps } from './components/IEmployeeListProps';

export interface IEmployeeListWebPartProps {
  description: string;
  listNameOrId: string;
  recordCount: number;
}

export default class EmployeeListWebPart extends BaseClientSideWebPart<IEmployeeListWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<IEmployeeListProps> = React.createElement(
      EmployeeList,
      {
        description: this.properties.description,
        listNameOrId: this.properties.listNameOrId || '7b906718-e108-42ac-92df-999e1c6e6e65',
        recordCount: this.properties.recordCount ?? 10,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        getAccessToken: async (): Promise<string> => {
          const provider = await this.context.aadTokenProviderFactory.getTokenProvider();
          return await provider.getToken('https://graph.microsoft.com');
        },
        spHttpClient: this.context.spHttpClient,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.all([
      this._getEnvironmentMessage(),
      this._fetchLists()
    ]).then(([message]) => {
      this._environmentMessage = message;
    });
  }

  private async _fetchLists(): Promise<void> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Id,Title&$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        this._listOptions = data.value.map((list: { Id: string; Title: string }) => ({
          key: list.Id,
          text: list.Title
        }));
      }
    } catch (e) {
      console.error('Failed to fetch lists', e);
    }
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
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
                }),
                PropertyPaneDropdown('listNameOrId', {
                  label: "Çalışanların Çekileceği Liste (SharePoint)",
                  options: this._listOptions,
                  selectedKey: this.properties.listNameOrId
                }),
                PropertyPaneSlider('recordCount', {
                  label: strings.RecordCountFieldLabel,
                  min: 1,
                  max: 100,
                  value: this.properties.recordCount !== undefined ? this.properties.recordCount : 10,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
