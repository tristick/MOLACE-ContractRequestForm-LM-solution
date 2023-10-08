import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMolaceContractRequestFormLmSolutionProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}

export interface ICustomer {

  Title:string;
  Reference:string;
  RefCode:string;
}




