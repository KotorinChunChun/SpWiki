declare interface ISpWikiCommandSetStrings {
  FullPageCommandTitle: string;
  PreviewCommandTitle: string;
  OpenInAppCommandTitle: string;
  DialogTitle: string;
  FullPageLabel: string;
  CloseLabel: string;
  CopyPathLabel: string;
  OpenInAppFailedMessage: string;
  LoadingMessage: string;
  NotFoundMessage: string;
  ErrorMessage: string;
  NotConfiguredMessage: string;
  HeadingCopyLabel: string;
  HeadingToggleLabel: string;
}

declare module 'SpWikiCommandSetStrings' {
  const strings: ISpWikiCommandSetStrings;
  export = strings;
}
