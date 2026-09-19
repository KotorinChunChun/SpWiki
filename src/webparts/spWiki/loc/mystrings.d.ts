declare interface ISpWikiWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  TitleFieldDescription: string;
  DefaultTitle: string;
  HomeButtonLabel: string;
  TagGroupName: string;
  ToolbarGroupName: string;

  LibraryNameFieldLabel: string;
  LibraryNameFieldDescription: string;
  StartFileFieldLabel: string;
  StartFileFieldDescription: string;
  SiteUrlFieldLabel: string;
  SiteUrlFieldDescription: string;

  TagFieldLabel: string;
  TagFieldLabel2: string;
  TagFieldLabel3: string;
  TagFieldNone: string;
  TagFieldNeedsLibrary: string;
  TagFieldLoading: string;
  TagFieldNoneFound: string;
  TagFieldReady: string;
  TagSearchPageLabel: string;
  TagSearchPageDescription: string;

  ShowTagEditorLabel: string;
  ShowEditButtonLabel: string;
  ShowOpenInAppButtonLabel: string;
  AppProtocolLabel: string;
  AppProtocolDescription: string;

  LoadingMessage: string;
  NotFoundMessage: string;
  ErrorMessage: string;
  LibraryNotFoundMessage: string;
  NotConfiguredMessage: string;
  PageLabel: string;

  TagEditLabel: string;
  EditButtonLabel: string;
  OpenInAppButtonLabel: string;


  TagPanelTitle: string;
  TagTextHint: string;
  TagNoChoices: string;
  TagFieldMissingMessage: string;
  TagSaveFailedMessage: string;

  TagAddLabel: string;
  TagAddPlaceholder: string;
  TagAddDisabledPlaceholder: string;
  TagAddHint: string;
  TagAddDisabledHint: string;

  ShowPermissionsButtonLabel: string;
  PermUnknownTitle: string;
  PermLockedTitle: string;
  PermUnlockedTitle: string;
  NotFoundReasonMessage: string;
  NotFoundTreeLabel: string;
  BackLabel: string;
  PermTitle: string;
  PermInheritedFromLabel: string;
  PermUniqueLabel: string;
  PermTeamLabel: string;
  PermFolderSuffix: string;
  PermSearchButton: string;
  PermSearchFailedMessage: string;
  PermWhoLabel: string;
  PermRestrictLabel: string;
  PermRestoreLabel: string;
  PermConfirmTitle: string;
  PermConfirmRestrict: string;
  PermConfirmRestore: string;
  PermConfirmYes: string;
  PermAddLabel: string;
  PermAddPlaceholder: string;
  PermAddButton: string;
  PermRemoveLabel: string;
  PermSearchingMessage: string;
  PermNoResultsMessage: string;
  PermBrokenTitle: string;
  PermBrokenIntro: string;
  PermFileScopeWarning: string;
  PermLoadingMessage: string;
  PermLoadFailedMessage: string;
  PermNoEffectMessage: string;
  PermChangeFailedMessage: string;
  SearchRestrictedOnlyLabel: string;
  PermLockedShortLabel: string;
  SearchLabel: string;
  SearchTitle: string;
  SearchTermPlaceholder: string;
  SearchModeName: string;
  SearchModeContent: string;
  SearchScopeLabel: string;
  SearchScopePlaceholder: string;
  SearchSubmitLabel: string;
  SearchBusyMessage: string;
  SearchProgressMessage: string;
  SearchResultCount: string;
  SearchEmptyMessage: string;
  SearchPromptMessage: string;
  HelpLabel: string;
  ChromeFullScreenLabel: string;
  ChromeSharePointLabel: string;
  HelpTitle: string;
  HelpDocumentsTitle: string;
  HelpSetupTitle: string;
  HelpBackLabel: string;
  HelpCommandSetIntro: string;
  HelpComponentId: string;
  HelpLocation: string;
  HelpListTemplate: string;
  HelpComponentProperties: string;
  HelpExtensionsList: string;
  CopyLabel: string;
  CopiedLabel: string;
  FolderPageEmptyMessage: string;
  OutlineButtonLabel: string;
  OutlineTitle: string;
  OutlineEmptyMessage: string;
  MenuButtonLabel: string;
  PathButtonLabel: string;
  NewPageTitle: string;
  NewPageNamePlaceholder: string;
  NewPageFormatLabel: string;
  NewPageOpenWithLabel: string;
  NewPageOpenInSharePoint: string;
  NewPageOpenInApp: string;
  CreateLabel: string;
  LocationTitle: string;
  RootLabel: string;
  FolderEmptyMessage: string;
  CreateFailedMessage: string;
  NewPageCreatedMessage: string;

  RenameTitle: string;
  RenameFolderLabel: string;
  RenameNameLabel: string;
  RenameHint: string;
  RenameFailedMessage: string;
  InvalidNameMessage: string;

  HeadingCopyLabel: string;
  HeadingToggleLabel: string;
  LinkCopiedMessage: string;
  CodeCopyLabel: string;
  CodeCopiedMessage: string;

  DirectiveEmptyMessage: string;

  TagSearchTitle: string;
  TagSearchEmpty: string;

  OpenInAppFailedMessage: string;
  SaveLabel: string;
  CancelLabel: string;
  CloseLabel: string;
}

declare module 'SpWikiWebPartStrings' {
  const strings: ISpWikiWebPartStrings;
  export = strings;
}
