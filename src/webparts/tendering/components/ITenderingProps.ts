import type { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITenderingProps {
  context: WebPartContext;
}

export interface ITenderRecord {
  Id: number;

  title: string;
  opportunityNumber: string;
  type: string;
  status: string;
  customer: string;
  deadline: string;
  manager: string;
  engineer: string;
  comment: string;
  domain: string;
  lastUpdated: string;

  handoverToTendering: boolean;
}