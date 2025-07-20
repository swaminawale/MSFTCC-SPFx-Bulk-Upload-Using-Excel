import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBulkUploadSpFxProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context: WebPartContext;
}

export interface ISharePointItem {
  Title: string;
  FirstName: string;
  LastName: string;
  WorkEmail: string;
  PersonalEmail: string;
  BirthDate: Date;
  HireDate: Date;
  WorkMode: string;
  Salary: number;
  IsMarried: boolean;
  SocialProfile: {
    Url: string;
  };
  JobTitle: string;
  About: string;
}
