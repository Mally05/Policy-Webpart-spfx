declare interface IPolicyWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  ListFieldLabel: string;
  Title:string;
  ListFieldLabel:string;
  SelectListLabel: string;
}

declare module 'PolicyWebPartStrings' {
  const strings: IPolicyWebPartStrings;
  export = strings;
}
