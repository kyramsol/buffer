function transferStoredChanges() {
  const properties = PropertiesService.getScriptProperties();
  QubGlobalLibrary.transferStoredChanges(properties);
}
