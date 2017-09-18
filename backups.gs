function backupSheetToXlsx() {
  //Backup this sheet to backups folder
  var lastBackupToXlsx = loadSetting('lastBackupToXlsx');
  if (!lastBackupToXlsx || lastBackupToXlsx !== formatDate(new Date())) {
    try {
      console.info("Backing up main spreadsheet...");
      console.time("Spreadsheet backup completed! Time: ");
      var ss = SpreadsheetApp.getActive();
      var autoBackupFolderId = loadSetting('autoBackupFolderId'); //Backup storage folder
      var folder = DriveApp.getFolderById(autoBackupFolderId);
      var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + ss.getId() + "&exportFormat=xlsx";
      var params = {
        method      : "get",
        headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true
      };
      
      var blob = UrlFetchApp.fetch(url, params).getBlob();
      var fileDate = dateFileName(new Date());
	  var backupName = ss.getName() + " - " + fileDate + ".xlsx";
      
      blob.setName(backupName);
      folder.createFile(blob);
      deleteOlderBackups();
      saveSetting('lastBackupToXlsx', formatDate(new Date()));
    } catch (f) {
      console.error("backupSheetToXlsx had an error: " + f.toString());
    } finally {
      console.info("Main spreadsheet backed up to Google Drive as '" + backupName + "'");
      console.timeEnd("Spreadsheet backup completed! Time: ");
    }
  } else {
    console.warn('Main spreadsheet has already been backed up today!');
  }
}



function deleteOlderBackups() {
  console.info("Checking for outdated backups...");
  var autoBackupFolderId = loadSetting('autoBackupFolderId'); //Backup storage folder
  var folder = DriveApp.getFolderById(autoBackupFolderId);
  var maxBackupsToKeep = loadSetting('maxBackupsToKeep'); //keep max of 30 files in this folder
  
  var files = folder.getFilesByType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  var numOfFiles = countXlsxFilesInBackupFolder();
  console.log("Number of files found in backup folder: " + numOfFiles + ". Max is : " + maxBackupsToKeep);
  
  //If too many files, delete older ones until all is good
  while (numOfFiles > maxBackupsToKeep) {
    //find oldest file and delete
    var oldestFileId = findOldestBackup();
    var oldestFileName = DriveApp.getFileById(oldestFileId).getName();
    Drive.Files.remove(oldestFileId);
    console.info("Deleted outdated backup: " + oldestFileName);
    numOfFiles = countXlsxFilesInBackupFolder();
  }
}


function countXlsxFilesInBackupFolder() {
  //Count number of files in backup folder
  var autoBackupFolderId = loadSetting('autoBackupFolderId'); //Backup storage folder
  var folder = DriveApp.getFolderById(autoBackupFolderId);
  var files = folder.getFilesByType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");  //xlsx files
  var numOfFiles = 0;
  while (files.hasNext()) {
    var file = files.next();
    numOfFiles++;
  }
  return numOfFiles;
}


function findOldestBackup() {
  //Find oldest excel file in target folder
  var autoBackupFolderId = loadSetting('autoBackupFolderId'); //Backup storage folder
  var folder = DriveApp.getFolderById(autoBackupFolderId);
  var arryFileDates,file,fileDate,files,
      oldestDate,oldestFileId,oldestFileName,objFilesByDate,objFilesByName;
  console.log('Looking for oldest excel backup file in folder: ' + folder);
  arryFileDates = [];
  objFilesByDate = {};
  files = folder.getFilesByType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  fileDate = "";

  while (files.hasNext()) {
    file = files.next();
    fileDate = file.getDateCreated();
    objFilesByDate[fileDate] = file.getId(); //Create an object of file names by file ID
    arryFileDates.push(file.getDateCreated());
  }
  
  arryFileDates.sort(function(a,b){return a-b});
  oldestDate = arryFileDates[0];
  oldestFileId = objFilesByDate[oldestDate];
  oldestFileName = DriveApp.getFileById(oldestFileId).getName();
  console.log('Oldest File is: ' + oldestFileName + '. Date: ' + oldestDate + ". ID: " + oldestFileId);
  
  return oldestFileId;
}