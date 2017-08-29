function sendEmail() {
  //Scheduled to run at 4-5AM every day
  var emailAddress = 'fsr@opex.com';
  var subject = "snapshot";
  var region = loadSetting('region');
  var warehouse = loadSetting('warehouse');
  var enableAutoEmailing = loadSetting('enableAutoEmailing');
  var message = region + ',' + warehouse + ',';
  var today = new Date();
  if(today.getDay() == 6 || today.getDay() == 0) {
    //do not run on weekends
  } else {
    if (enableAutoEmailing) {
      MailApp.sendEmail(emailAddress, subject, message);
      console.log('Email Sent! ... To: ' + emailAddress + ' ... Subject: ' + subject + ' ... Message: ' + message);
    } else {
      console.log('Auto Emailing is disabled in settings!');
    }
  };
}

function updateSnapshot() {
  console.log('Starting updateSnapshot()');
  console.time("updateSnapshot time");
  
  var gmailSnapshotLabel = loadSetting('gmailSnapshotLabel');
  var warehouse = loadSetting('warehouse');
  var snapshotFolderId = loadSetting('snapshotFolderId');
  var latestSnapshotSheet = loadSetting('latestSnapshotSheet');
  
  var label = GmailApp.getUserLabelByName(gmailSnapshotLabel);
  var threads = label.getThreads();
  var msg = threads[0].getMessages();
  var attachment = msg[msg.length - 1].getAttachments();
  var blob = attachment[0].copyBlob();
  
  var date = msg[msg.length - 1].getDate();
  var NewXlsFileName = warehouse + ' ' + dateFileName(date) + '.xls';
  var folder = DriveApp.getFolderById(snapshotFolderId); //Snapshot Folder on gDrive
  
  GmailApp.markMessageRead(msg[0]);
  threads[0].moveToArchive();
  
  
  //Check if excel file with name exists:
  var NewXlsFileNameChk = folder.getFilesByName(NewXlsFileName);
  if (NewXlsFileNameChk.hasNext()) {
    console.log('Filename ' + NewXlsFileName + ' already exists, skipping writing of file.');
    //do nothing
  } else {
    console.log('File ' + NewXlsFileName + ' has been written to ' + folder);
    blob.setName(NewXlsFileName);
    folder.createFile(blob);
  }
  
  var excelFile = DriveApp.getFileById(findNewestFileId(snapshotFolderId));
  var gSheetName = folder.getFilesByName(latestSnapshotSheet);
  var resource = {
    title: latestSnapshotSheet,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{id: folder.getId()}],
  };
  
  //Check if gSheet with name already exists:
  if (gSheetName.hasNext()) {
    console.log(".gSheet file already exists, overwriting existing data.");
    var file = gSheetName.next();
    Drive.Files.update(resource, file.getId(), excelFile);
  } else {
    console.log("Existing .gSheet file not found, creating new file.");
    Drive.Files.insert(resource, excelFile);
  }
  
  copySnapshotData();  
  
  updateAllHelperLinks();
  
  console.log('Completed updateSnapshot()');
  console.timeEnd("updateSnapshot time");
}


function findNewestFileId(folderId) {
  //Find newest excel file in target folder
  var arryFileDates,file,fileDate,files,folder,folders,
      newestDate,newestFileID,objFilesByDate;
  console.log('Looking for newest excel file in FolderId: ' + folderId);
  folders = DriveApp.getFolderById(folderId);  
  arryFileDates = [];
  objFilesByDate = {};
  folder = folders;
  files = folder.getFilesByType("application/vnd.ms-excel");
  fileDate = "";

  while (files.hasNext()) {
    file = files.next();
    fileDate = file.getLastUpdated();
    objFilesByDate[fileDate] = file.getId(); //Create an object of file names by file ID
    arryFileDates.push(file.getLastUpdated());
  }
  arryFileDates.sort(function(a,b){return b-a});
  newestDate = arryFileDates[0];
  newestFileID = objFilesByDate[newestDate];
  console.log('Newest File ID: ' + newestFileID + '\n Newest File Date: ' + newestDate);
  return newestFileID;
};


function copySnapshotData() {
  console.log('Copying data from temp sheet...');
  var snapshotSSid = loadSetting('snapshotSSid');
  var mainSSid = loadSetting('mainSSid');
  var sheetNameSnapshot = loadSetting('sheetNameSnapshot');
  var snapshotSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameSnapshot);
  
  var sss = SpreadsheetApp.openById(snapshotSSid); // source
  var ss = sss.getActiveSheet();
  var SRange = ss.getDataRange();
  var A1Range = SRange.getA1Notation();
  var SData = SRange.getValues();
  
  var tss = SpreadsheetApp.openById(mainSSid); // target
  var ts = snapshotSheet;
  ts.clear({contentsOnly: true}); // Clear the Google Sheet before copy
  ts.getRange(A1Range).setValues(SData);
  console.log('Data copied to Snapshot sheet');
};