declare interface IExStorageWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  // 1 - Analytics options
  analyticsWeb: string;
  analyticsList: string;
  analyticsListErrors: string;
  
}

declare module 'ExStorageWebPartStrings' {
  const strings: IExStorageWebPartStrings;
  export = strings;
}
