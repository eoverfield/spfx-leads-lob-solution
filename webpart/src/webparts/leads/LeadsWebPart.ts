import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButtonType,
  PropertyPaneButton,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from "@microsoft/sp-property-pane";

import * as strings from 'LeadsWebPartStrings';
import { Leads, ILeadsProps, LeadView } from './components/Leads';
import { HttpClientResponse, HttpClient, MSGraphClientV3 } from '@microsoft/sp-http';
import { ILeadsSettings, LeadsSettings } from '../../LeadsSettings';

export interface ILeadsWebPartProps {
  description: string;
  demo: boolean;
  region: string;
  quarterlyOnly: boolean;
}

export default class LeadsWebPart extends BaseClientSideWebPart<ILeadsWebPartProps> {
  private _needsConfiguration: boolean;
  private _leadsApiUrl: string;
  private _connectionStatus: string;
  private _queryParameters: URLSearchParams;
  private _view?: LeadView;
  private _msGraphClient: MSGraphClientV3;
  private _settings?: ILeadsSettings;

  protected onInit(): Promise<void> {
    return this.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3): Promise<ILeadsSettings> => {
        this._msGraphClient = client;

        LeadsSettings.initialize(this._msGraphClient, this.context.httpClient);
        return LeadsSettings.getSettings();
      })
      .then((settings: ILeadsSettings): Promise<void> => {
        this._settings = settings;

        if (this.properties.demo) {
          this._needsConfiguration = false;
          return Promise.resolve();
        }

        return this._getApiUrl();
      });
  }

  private _getLeadView(): LeadView | undefined {
    const view: string = this._queryParameters.get('view');
    const supportedViews: string[] = ['new', 'mostProbable', 'recentComments', 'requireAttention'];

    if (!view || supportedViews.indexOf(view) < 0) {
      return undefined;
    }

    return LeadView[view];
  }

  public render(): void {
    const element: React.ReactElement<ILeadsProps> = React.createElement(
      Leads,
      {
        demo: this.properties.demo,
        httpClient: this.context.httpClient,
        // eslint-disable-next-line
        host: (this.context as any)._host,
        leadsApiUrl: this._leadsApiUrl,
        msGraphClient: this._msGraphClient,
        teamsContext: this.context.sdks.microsoftTeams,
        needsConfiguration: this._needsConfiguration,
        view: this._view
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _testConnection = (): void => {
    this._connectionStatus = 'Connecting to the API...';
    this.context.propertyPane.refresh();

    this.context.httpClient
      .get(this._leadsApiUrl, HttpClient.configurations.v1)
      .then((res: HttpClientResponse): void => {
        this._connectionStatus = 'Connection OK';
        this.context.propertyPane.refresh();
      }, (error): void => {
        this._connectionStatus = `Connection error: ${error}`;
        this.context.propertyPane.refresh();
      });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const config: IPropertyPaneConfiguration = {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "Data",
              groupFields: [
                PropertyPaneToggle('demo', {
                  label: 'Demo mode',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            },
            {
              groupName: "Filtering options",
              isCollapsed: false,
              groupFields: [
                PropertyPaneDropdown('region', {
                  label: 'Region',
                  options: [
                    { key: '1', text: 'America' },
                    { key: '2', text: 'Europe' },
                    { key: '3', text: 'Asia' },
                    { key: '4', text: 'South Pole' }
                  ]
                }),
                PropertyPaneToggle('quarterlyOnly', {
                  label: 'Only this quarter',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            },
            {
              groupName: "Connection Validation",
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('apiUrl', {
                  label: 'LOB API URL',
                  placeholder: 'Not configured',
                  value: this._leadsApiUrl,
                  disabled: true
                }),
                PropertyPaneLabel('spacer1', { text: '' }),
                PropertyPaneButton('testConnection', {
                  buttonType: PropertyPaneButtonType.Primary,
                  disabled: this.properties.demo || this._needsConfiguration,
                  onClick: this._testConnection,
                  text: 'Test connection'
                }),
                PropertyPaneLabel('connectionStatus', {
                  text: this._needsConfiguration ? 'Required tenant property LeadsApiUrl not set' : this._connectionStatus
                })
              ]
            }
          ]
        }
      ]
    };
    return config;
  }

  // eslint-disable-next-line
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue, newValue): void {
    if (propertyPath === 'demo') {
      if (newValue === true) {
        this._needsConfiguration = false;
      }
      else {
        this._needsConfiguration = true;
        // eslint-disable-next-line
        this._getApiUrl(true);
      }
    }
  }

  private _getApiUrl(reRender: boolean = false): Promise<void> {
    if (this._leadsApiUrl) {
      this._needsConfiguration = false;
      if (reRender) {
        this.render();
      }
      return Promise.resolve();
    }

    return new Promise<void>((resolve: () => void, reject: (err) => void): void => {
      LeadsSettings
        .getLeadsApiUrl(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl)
        .then((leadsApiUrl: string): void => {
          this._leadsApiUrl = leadsApiUrl;
          this._needsConfiguration = !this._leadsApiUrl;
          if (reRender) {
            this.render();
          }
          resolve();
        }, err => resolve());
    });
  }

  // eslint-disable-next-line
  protected onAfterDeserialize(deserializedObject: any, dataVersion: Version): ILeadsWebPartProps {
    const props: ILeadsWebPartProps = deserializedObject;
    this._queryParameters = new URLSearchParams(document.location.search);
    this._view = this._getLeadView();

    if (this.context.sdks.microsoftTeams && typeof this._view !== 'undefined') {
      if (!this._settings) {
        this._settings = {
          demo: true,
          quarterlyOnly: true,
          region: ""
        };
      }

      props.demo = this._settings.demo;
      props.quarterlyOnly = this._settings.quarterlyOnly;
      props.region = this._settings.region;
    }

    return props;
  }
}
