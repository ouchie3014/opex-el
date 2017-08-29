function initSheet() {
  //Run this only if any of these defaults are changed. This will overwrite all settings in scriptProperties.
  
  //Sheet Names:
  var sheetNameMainDefault = "MainDB";
  var sheetNameInventoryDefault = "InventoryDB";
  var sheetNamePMDefault = "PM Summary";
  var sheetNameSnapshotDefault = "Snapshot";
  
  //Email Settings:
  var enableAutoEmailingDefault = true;
  var warehouseDefault = 'zMHC600004';
  var regionDefault = '09';
  var gmailSnapshotLabelDefault = 'Snapshot';
  var latestSnapshotSheetDefault = 'Latest Snapshot';
  
  var emailSnapshotRequestHrDefault = 8; //(24hr) Time of day to request snapshot. Default is 8am.
  var updateSnapshotHrDefault = 9; //(24hr) Time of day to update snapshot. Recommended to be 1 hour after emailSnapshotRequestHr.
  
  var maxBackupsToKeepDefault = 30; //number of main database backups to keep
  
  //Location ID's:
  var mainSSidDefault = '1D-WzKPQQ9cUB_pe7pmHRp98Xt6TAbuLHUAdZ5PsIbjs';
  var snapshotSSidDefault = '171Wge_3u3qC_rKSxCnp3pPsvLGhd43EILxk6vXfBA7M';
  var snapshotFolderIdDefault = '0B6qhI9tGCNQzM0xRbmFSMTdfSVU';
  var autoBackupFolderIdDefault = '0BwtCesct9-1EUG9kX0liQ0gwRkk';
  
  
  // ----------------------------------------------------------------------------------------------- //
  // --------------------------------- DO NOT EDIT BELOW THIS LINE --------------------------------- //
  // ----------------------------------------------------------------------------------------------- //
  
  //Delete previous settings
  deleteAllSettings();
  
  //Store new settings
  var scriptProperties = PropertiesService.getScriptProperties();

  scriptProperties.setProperty('sheetNameMain', sheetNameMainDefault);
  scriptProperties.setProperty('sheetNameInventory', sheetNameInventoryDefault);
  scriptProperties.setProperty('sheetNamePM', sheetNamePMDefault);
  scriptProperties.setProperty('sheetNameSnapshot', sheetNameSnapshotDefault);

  scriptProperties.setProperty('enableAutoEmailing', enableAutoEmailingDefault);
  scriptProperties.setProperty('warehouse', warehouseDefault);
  scriptProperties.setProperty('region', regionDefault);
  scriptProperties.setProperty('gmailSnapshotLabel', gmailSnapshotLabelDefault);
  scriptProperties.setProperty('latestSnapshotSheet', latestSnapshotSheetDefault);
  
  scriptProperties.setProperty('emailSnapshotRequestHr', emailSnapshotRequestHrDefault);
  scriptProperties.setProperty('updateSnapshotHr', updateSnapshotHrDefault);
  
  scriptProperties.setProperty('maxBackupsToKeep', maxBackupsToKeepDefault);
  
  scriptProperties.setProperty('mainSSid', mainSSidDefault);
  scriptProperties.setProperty('snapshotSSid', snapshotSSidDefault);
  scriptProperties.setProperty('snapshotFolderId', snapshotFolderIdDefault);
  scriptProperties.setProperty('autoBackupFolderId', autoBackupFolderIdDefault);
  
  console.log("Sheet Initialized! Default settings stored!");
}


//Used to quickly get stored settings
function loadSetting(name) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var output = scriptProperties.getProperty(name) || '';
  return output;
}



function unused() {
//Location ID's:
  var mainSSid = scriptProperties.getProperty('mainSSid') || mainSSidDefault;
  var snapshotSSid = scriptProperties.getProperty('snapshotSSid') || snapshotSSidDefault;
  var snapshotFolderId = scriptProperties.getProperty('snapshotFolderId') || snapshotFolderIdDefault;
  
  
  //Sheet Names:
  var sheetNameMain = scriptProperties.getProperty('sheetNameMain') || sheetNameMainDefault;
  var sheetNameInventory = scriptProperties.getProperty('sheetNameInventory') || sheetNameInventoryDefault;
  var sheetNamePM = scriptProperties.getProperty('sheetNamePM') || sheetNamePMDefault;
  var sheetNameSnapshot = scriptProperties.getProperty('sheetNameSnapshot') || sheetNameSnapshotDefault;
  
  
  //Email Settings:
  var enableAutoEmailing = scriptProperties.getProperty('enableAutoEmailing') || enableAutoEmailingDefault;
  var gmailSnapshotLabel = scriptProperties.getProperty('gmailSnapshotLabel') || gmailSnapshotLabelDefault;
  var latestSnapshotSheet = scriptProperties.getProperty('latestSnapshotSheet') || latestSnapshotSheetDefault;
  var warehouse = scriptProperties.getProperty('warehouse') || warehouseDefault;
  var region = scriptProperties.getProperty('region') || regionDefault;
  
  var emailSnapshotRequestHr = scriptProperties.getProperty('emailSnapshotRequestHr') || emailSnapshotRequestHrDefault; //(24hr) Time of day to request snapshot. Default is 8am.
  var updateSnapshotHr = scriptProperties.getProperty('updateSnapshotHr') || updateSnapshotHrDefault; //(24hr) Time of day to update snapshot. Recommended to be 1 hour after emailSnapshotRequestHr.
  
  
  //Global variables:
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameMain);
  var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameInventory);
  var pmDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNamePM);
  var snapshotSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameSnapshot);
}




function saveSidebarSettings(form) {
  var scriptProperties = PropertiesService.getScriptProperties();

  scriptProperties.setProperty('sheetNameMain', form.sheetNameMain);
  scriptProperties.setProperty('sheetNameInventory', form.sheetNameInventory);
  scriptProperties.setProperty('sheetNamePM', form.sheetNamePM);
  scriptProperties.setProperty('sheetNameSnapshot', form.sheetNameSnapshot);

  scriptProperties.setProperty('enableAutoEmail', form.enableAutoEmail);
  scriptProperties.setProperty('warehouse', form.warehouse);
  scriptProperties.setProperty('region', form.region);
  scriptProperties.setProperty('gmailSnapshotLabel', form.gmailSnapshotLabel);
  scriptProperties.setProperty('latestSnapshotSheet', form.latestSnapshotSheet);
  
  scriptProperties.setProperty('emailSnapshotRequestHr', form.emailSnapshotRequestHr);
  scriptProperties.setProperty('updateSnapshotHr', form.updateSnapshotHr);
  
  scriptProperties.setProperty('maxBackupsToKeep', form.maxBackupsToKeep);
  
  scriptProperties.setProperty('mainSSid', form.mainSSid);
  scriptProperties.setProperty('snapshotSSid', form.snapshotSSid);
  scriptProperties.setProperty('snapshotFolderId', form.snapshotFolderId);
  scriptProperties.setProperty('autoBackupFolderId', form.autoBackupFolderId);
  
  console.log("Settings saved.");
  
}


function getPreferences() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetSettings = {
    sheetNameMain: scriptProperties.getProperty('sheetNameMain'),
    sheetNameInventory: scriptProperties.getProperty('sheetNameInventory'),
    sheetNamePM: scriptProperties.getProperty('sheetNamePM'),
    sheetNameSnapshot: scriptProperties.getProperty('sheetNameSnapshot'),
    
    enableAutoEmailing: scriptProperties.getProperty('enableAutoEmailing'),
    warehouse: scriptProperties.getProperty('warehouse'),
    region: scriptProperties.getProperty('region'),
    gmailSnapshotLabel: scriptProperties.getProperty('gmailSnapshotLabel'),
    latestSnapshotSheet: scriptProperties.getProperty('latestSnapshotSheet'),
    
    emailSnapshotRequestHr: scriptProperties.getProperty('emailSnapshotRequestHr'),
    updateSnapshotHr: scriptProperties.getProperty('updateSnapshotHr'),
    
    maxBackupsToKeep: scriptProperties.getProperty('maxBackupsToKeep'),
    
    mainSSid: scriptProperties.getProperty('mainSSid'),
    snapshotSSid: scriptProperties.getProperty('snapshotSSid'),
    snapshotFolderId: scriptProperties.getProperty('snapshotFolderId'),
    autoBackupFolderId: scriptProperties.getProperty('autoBackupFolderId')
  };
  console.log("Settings loaded.");
  return sheetSettings;
}

function restoreDefaultSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();

  scriptProperties.setProperty('sheetNameMain', sheetNameMainDefault);
  scriptProperties.setProperty('sheetNameInventory', sheetNameInventoryDefault);
  scriptProperties.setProperty('sheetNamePM', sheetNamePMDefault);
  scriptProperties.setProperty('sheetNameSnapshot', sheetNameSnapshotDefault);

  scriptProperties.setProperty('enableAutoEmailing', enableAutoEmailingDefault);
  scriptProperties.setProperty('warehouse', warehouseDefault);
  scriptProperties.setProperty('region', regionDefault);
  scriptProperties.setProperty('gmailSnapshotLabel', gmailSnapshotLabelDefault);
  scriptProperties.setProperty('latestSnapshotSheet', latestSnapshotSheetDefault);
  
  scriptProperties.setProperty('emailSnapshotRequestHr', emailSnapshotRequestHrDefault);
  scriptProperties.setProperty('updateSnapshotHr', updateSnapshotHrDefault);
  
  scriptProperties.setProperty('maxBackupsToKeep', maxBackupsToKeepDefault);
  
  scriptProperties.setProperty('mainSSid', mainSSidDefault);
  scriptProperties.setProperty('snapshotSSid', snapshotSSidDefault);
  scriptProperties.setProperty('snapshotFolderId', snapshotFolderIdDefault);
  scriptProperties.setProperty('autoBackupFolderId', autoBackupFolderIdDefault);
  
  console.log("Settings restored to default values.");
  
}

function deleteAllSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
  console.log("All stored settings deleted!.");
}