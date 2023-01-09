declare interface IWebPartTabsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'WebPartTabsWebPartStrings' {
  const strings: IWebPartTabsWebPartStrings;
  export = strings;
}
