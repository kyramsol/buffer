function saveWorkingTabActions(e) {
  const locker = LockService.getScriptLock()
  locker.waitLock(30000);
  {
    const properties = PropertiesService.getScriptProperties();
    QubGlobalLibrary.saveWorkingTabActions(e, properties);
  }
  locker.releaseLock();
}