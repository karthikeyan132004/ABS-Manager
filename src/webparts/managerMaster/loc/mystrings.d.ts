declare interface IManagerMasterWebPartStrings {
  PropertyPaneDescription: string;
  GeneralGroupName: string;
  QuotationGroupName: string;
  LeadsGroupName: string;
  ActivityGroupName: string;
  ProjectsGroupName: string;
  TitleFieldLabel: string;
  DefaultDashboardTitle: string;
  ItemLimitFieldLabel: string;
  ProbabilityThresholdFieldLabel: string;
  ListTitleFieldLabel: string;
  ColumnsFieldLabel: string;
  ColumnsFieldDescription: string;
  DateFieldLabel: string;
  DateFieldDescription: string;
  StatusFieldLabel: string;
  StatusFieldDescription: string;
  QuotationStatusFieldDescription: string;
  AmountFieldLabel: string;
  RemarksFieldLabel: string;
  OptionalFieldDescription: string;
  LeadLinkFieldLabel: string;
  LeadLinkFieldDescription: string;
  LeadStageFieldLabel: string;
  LeadStageFieldDescription: string;
  ProbabilityFieldLabel: string;
  ProbabilityFieldDescription: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'ManagerMasterWebPartStrings' {
  const strings: IManagerMasterWebPartStrings;
  export = strings;
}
