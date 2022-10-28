declare interface IPolicyWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ListFieldLabel: string;
  Title:string;
  ListFieldLabel:string;
  SelectListLabel: string;
  SiteCollection:string;
  CheckboxLabel: string;
  Accepted: string
  On: string
  LoadingMessage: string
  CheckboxPlaceholder:string;
  Description:string
  TitlePlaceholder: string;
}

declare module 'PolicyWebPartStrings' {
  const strings: IPolicyWebPartStrings;
  export = strings;
}
