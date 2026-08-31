import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICrmProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  context: WebPartContext;
}