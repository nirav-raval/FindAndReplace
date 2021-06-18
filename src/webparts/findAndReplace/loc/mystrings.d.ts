declare interface IFindAndReplaceWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'FindAndReplaceWebPartStrings' {
  const strings: IFindAndReplaceWebPartStrings;
  export = strings;
}
