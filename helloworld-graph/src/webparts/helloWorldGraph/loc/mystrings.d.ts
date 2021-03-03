declare interface IHelloWorldGraphWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'HelloWorldGraphWebPartStrings' {
  const strings: IHelloWorldGraphWebPartStrings;
  export = strings;
}
