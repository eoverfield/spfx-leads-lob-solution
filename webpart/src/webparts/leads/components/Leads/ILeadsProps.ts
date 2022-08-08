import { HttpClient, MSGraphClientV3 } from "@microsoft/sp-http";
import { IMicrosoftTeams } from "@microsoft/sp-webpart-base";
import { LeadView } from "..";

export interface ILeadsProps {
  demo: boolean;
  httpClient: HttpClient;
  // eslint-disable-next-line
  host?: any;
  leadsApiUrl: string;
  msGraphClient: MSGraphClientV3;
  needsConfiguration: boolean;
  teamsContext?: IMicrosoftTeams;
  view?: LeadView;
}
