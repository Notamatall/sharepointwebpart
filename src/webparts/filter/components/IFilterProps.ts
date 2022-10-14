import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFilterProps {
  context: WebPartContext;
  list: any;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
