import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IEmployeeEvaluationProps {
  sp: SPFI;
  graphClient: MSGraphClientV3;
  siteUrl: string;
  context: WebPartContext;
}

export type IUser = {
  id: string;
  displayName: string;
  userPrincipalName: string;
  secondaryText: string;
  employeeType?: string;      // סוג עובד
  department?: string;        // מחלקה
  subDepartment?: string;     // תת-מחלקה
  directManager?: string;     // מנהל ישיר
  indirectManager?: string;   // מנהל עקיף
  employeeId : string;
};

export type IGroup = {
  id: string;
  displayName: string;
};
