declare interface IEcStorageWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  // 1 - Analytics options
  analyticsWeb: string;
  analyticsList: string;
  
}

declare module 'EcStorageWebPartStrings' {
  const strings: IEcStorageWebPartStrings;
  export = strings;
}
