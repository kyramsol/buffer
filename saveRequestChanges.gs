function saveRequestsActions() {
  const locker = LockService.getScriptLock()
  locker.waitLock(30000);
  {
    const properties = PropertiesService.getScriptProperties();
    QubGlobalLibrary.saveRequestsActions(properties)
  }
  locker.releaseLock();
}
