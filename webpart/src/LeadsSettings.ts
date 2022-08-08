import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3, HttpClient, HttpClientResponse } from "@microsoft/sp-http";

export interface ILeadsSettings {
  demo: boolean;
  quarterlyOnly: boolean;
  region?: string;
}

export class LeadsSettings {
  private static _settingsFileName: string = 'LeadsSettings.json';
  private static _settingsFileUrl: string = `/me/drive/special/approot:/${LeadsSettings._settingsFileName}`;
  private static _graphClient: MSGraphClientV3;
  private static _httpClient: HttpClient;

  public static initialize(graphHttpClient: MSGraphClientV3, httpClient: HttpClient): void {
    this._graphClient = graphHttpClient;
    this._httpClient = httpClient;
  }

  public static getSettings(): Promise<ILeadsSettings> {
    if (!this._graphClient) {
      throw new Error('Initialize LeadsSettings before managing settings');
    }

    const defaultSettings: ILeadsSettings = {
      demo: true,
      quarterlyOnly: true
    };

    return this._graphClient
      .api(`${LeadsSettings._settingsFileUrl}?select=@microsoft.graph.downloadUrl`)
      .get()
      .then((response: { '@microsoft.graph.downloadUrl': string }): Promise<HttpClientResponse> => {
        return this._httpClient
          .get(response['@microsoft.graph.downloadUrl'], HttpClient.configurations.v1);
      })
      .then((response: HttpClientResponse): Promise<string> => {
        if (response.ok) {
          return response.text();
        }

        return Promise.reject(response.statusText);
      })
      .then((settingsString: string): Promise<ILeadsSettings> => {
        try {
          const settings: ILeadsSettings = JSON.parse(settingsString);
          return Promise.resolve(settings);
        }
        catch (e) {
          return Promise.resolve(defaultSettings);
        }
      }, reject => Promise.resolve(defaultSettings));
  }

  public static setSettings(settings: ILeadsSettings): Promise<void> {
    if (!this._graphClient) {
      throw new Error('Initialize LeadsSettings before managing settings');
    }

    return this._graphClient
      .api(`${LeadsSettings._settingsFileUrl}:/content`)
      .header('content-type', 'text/plain')
      .put(JSON.stringify(settings));
  }

  public static getLeadsApiUrl(spHttpClient: SPHttpClient, siteUrl: string): Promise<string> {
    return new Promise<string>((resolve: (leadsApiUrl: string) => void, reject: (error) => void): void => {
      spHttpClient
        .get(`${siteUrl}/_api/web/GetStorageEntity('LeadsApiUrl')`, SPHttpClient.configurations.v1)
        .then((res: SPHttpClientResponse) => {
          if (!res.ok) {
            return reject(res.statusText);
          }

          return res.json();
        })
        .then((property: { Value?: string }) => {
          if (property.Value) {
            resolve(property.Value);
          }
          else {
            reject('Property not found');
          }
        }, (error): void => {
          reject(error);
        });
    });
  }
}