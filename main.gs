//Add menu items:
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OPEX Scripts')
      .addItem('Update HTML Links', 'updateAllHelperLinks')
      .addItem('Request Snapshot Via E-mail', 'sendEmail')
      .addItem('Update Snapshot From Drive', 'updateSnapshot')
      .addItem('Settings', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('OPEX Settings')
      .setWidth(335);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function updateAllTriggers() {
  console.info('Updating all triggers...');
  var emailSnapshotRequestHr = loadSetting('emailSnapshotRequestHr');
  var updateSnapshotHr = loadSetting('updateSnapshotHr');
  var lastTriggerUpdate = loadSetting('lastTriggerUpdate');

  if (!lastTriggerUpdate || lastTriggerUpdate !== formatDate(new Date())) {
    //Delete all existing triggers
    deleteTriggers();
    
    // Trigger every Monday at 09:00.
    ScriptApp.newTrigger('sendEmail')
        .timeBased()
        .atHour(emailSnapshotRequestHr)
        .everyDays(1)
        .create();
  
    ScriptApp.newTrigger('updateSnapshot')
        .timeBased()
        .atHour(updateSnapshotHr)
        .everyDays(1)
        .create();
  
    ScriptApp.newTrigger('backupSheetToXlsx')
        .timeBased()
        .atHour(1)
        .everyDays(1)
        .create();
  
    ScriptApp.newTrigger('monitorRecentlySubmitted')
        .timeBased()
        .everyMinutes(1)
        .create();
  
    ScriptApp.newTrigger('updateAllTriggers')
        .timeBased()
        .atHour(1)
        .everyDays(1)
        .create();
    
    saveSetting('lastTriggerUpdate', formatDate(new Date()));
    console.info("All triggers updated");
  } else {
    console.warn('All triggers have already been updated today!');
  }
  
}

function deleteTriggers() {
  //Delete all triggers
  var allTriggers = ScriptApp.getProjectTriggers();  
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  console.info("All triggers deleted");
}

//Formats date into DD/MM/YYYY
function formatDate(rawDate) {
  var date = "";
  if (rawDate) {
    date = Utilities.formatDate(rawDate, "GMT", "MM/dd/yyyy");
  }
  return date;
}

//Returns date with format yyyy-MM-DD
function dateFileName(rawDate) {
  var date = "";
  if (rawDate) {
    date = Utilities.formatDate(rawDate, "GMT", "yyyy-MM-dd");
  }
  return date;
}

function onEdit () {
  var sheetNameMain = loadSetting('sheetNameMain');
  var activeRow = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getRow();
  console.info("Sheet has been edited, updating row: " + activeRow);
  updateHelperLinks(activeRow, 1); //Update last edited row
}

function getMainData (startRow,numRows) {
  var sheetNameMain = loadSetting('sheetNameMain');
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameMain);
  var lastRow = mainSheet.getLastRow();
  
  var data = mainSheet.getRange(startRow,4,numRows,7).getValues();
  console.log("Downloaded MainDB sheet data");
  return data;
}
function getInvData () {
  var sheetNameInventory = loadSetting('sheetNameInventory');
  var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameInventory);
  var data = inventorySheet.getDataRange().getValues();
  console.log("Downloaded InventoryDB sheet data");
  return data;
}

function getPmData () {
  var sheetNamePM = loadSetting('sheetNamePM');
  var pmDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNamePM);
  var data = pmDataSheet.getDataRange().getValues();
  console.log("Downloaded Maintenance sheet data");
  return data;
}

function updateHelperLinks(startRow,numRows) {
  console.info("Updating helper links...");
  console.time("updateHelperLinks time");
  console.log("Start row: " + startRow + ". Number of rows: " + numRows);
  
  var sheetNameMain = loadSetting('sheetNameMain');
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameMain);
  var lastRow = mainSheet.getLastRow();
    
  if (!startRow) {
    startRow = 2;
  } else if (startRow < 2) {
    startRow = 2;
  } else if (startRow > lastRow) {
    startRow = 2;
  }
  
  if (!numRows) {
    numRows = lastRow-startRow+1;
  } else if (numRows < 1) {
    numRows = 1;
  } else if (numRows > (lastRow-startRow+1)) {
    numRows = lastRow-startRow+1;
  }

  var mainRange = getMainData(startRow,numRows);
  var inventoryRange = getInvData();
  var pmData = getPmData();
  
  var partsUsed = [];
  var partsRep = [];
  var serialNumbers = [];
  for (var i=0, iLen=mainRange.length; i<iLen; i++) {
    serialNumbers[i] = mainRange[i][0];
    partsUsed[i] = mainRange[i][5];
    partsRep[i] = mainRange[i][6];
  }

  var serialNumberLinks = snTooltip(serialNumbers,pmData);
  var partsUsedLinks = partsTooltip(partsUsed,inventoryRange);
  var partsRepLinks = partsTooltip(partsRep,inventoryRange);

  mainSheet.getRange(startRow, 14, serialNumberLinks.length, 1).setValues(serialNumberLinks);
  mainSheet.getRange(startRow, 15, partsUsedLinks.length, 1).setValues(partsUsedLinks);
  mainSheet.getRange(startRow, 16, partsRepLinks.length, 1).setValues(partsRepLinks);

  console.timeEnd("updateHelperLinks time");
  console.info("All helper links updated.");
}



function partsTooltip (parts,inventory) {
  console.log("Running partsTooltip()");
  console.time("partsTooltip time");
  
  if (typeof inventory === 'undefined') {
    console.log("Inventory argument is undefined, downloading inventory data");
    inventory = getInvData();
  }
  
  for (var k=1, kLen=inventory.length; k<kLen; k++) {
    if (inventory[k][0] == "") {
      //do nothing
    } else {
      var partNum = inventory[k][0].toString();
      var partDescRaw = inventory[k][1];
      var partDesc = partDescRaw.replace(/'/g, "&apos;").replace(/"/g, "&quot;");
      var actualQty = '\n Actual Qty: ' + inventory[k][5];
      var snapQty = '\n Snapshot Qty: ' + inventory[k][6];
      if (!inventory[k][11]) {
        var snapDate = '';
      } else {
        var snapDate = ' (as of: ' + formatDate(inventory[k][11]) + ')';
      }      
      var highDollar = '\n High Dollar Item: ' + inventory[k][10];
      var minQty = '\n Min: ' + inventory[k][7];
      var maxQty = '\n Max: ' + inventory[k][8];
      
      var lookFor = new RegExp(partNum,'g');
      var replaceWith = '<a href="Inventory.html?part=' + partNum + '" class="partTooltip" title="' + partDesc + actualQty + snapQty + snapDate + minQty + maxQty + highDollar +'" target="_blank">' + partNum + '</a>';

      for (var j=0, jLen=parts.length; j<jLen; j++) {
        if (parts[j] == "") {
          //do nothing if entry is empty
        } else {
          parts[j] = parts[j].toString().replace(lookFor, replaceWith);
        }
      }
    }
  }  
  var output = [];
  for (var k=0, kLen=parts.length; k<kLen; k++) {
    output.push([parts[k]]);
  }
  
  console.timeEnd("partsTooltip time");
  console.log("Completed partsTooltip()");
  return output;
}


function snTooltip(sn,pm) {
  console.log("Running snTooltip()");
  console.time("snTooltip time");
  
  if (typeof pm === 'undefined') {
    console.log("PM argument is undefined, downloading PM data"); 
    pm = getPmData();
  }
  
  var output = [];
  
  for (var k=0, kLen=sn.length; k<kLen; k++) {
    var pmType = "";
    var pmLastRaw = "";
    var pmName = "";
    var pmNextRaw = "";
    var pmDays = "";
    var row = sn[k];
    
    if (sn[k] == "") {
      //return empty string if serial number is not present, may be redundant
      row = "";
    } else if (sn[k].substring(0,2) == "PP") {
      //if first two characters in SN are "PP" then it's an aisle not a bot
      
      var increment = 0;
      for (var x=1, xLen=pm.length; x<xLen; x++) {
        
        if (pm[x][0].substring(0,5) == sn[k]) {
          //there are 3 matches for each aisle, and for each match, there are 4 pieces of data
          //increment is the 3 matches, pm[x][3-6] are the 4 pieces of data
          
          if (increment == 0) {
            //weekly
            if (pm[x][3] == "Not Done") {
              var pmWeeklyLast = 'Weekly PM:\n A weekly PM has not been done yet! \n\n';
              var pmWeeklyName = '';
              var pmWeeklyNext = '';
              var pmWeeklyDays = '';
            } else {
              var pmWeeklyLast = 'Weekly PM:\n Last: ' + formatDate(pm[x][3]);
              var pmWeeklyName = ' - ' + pm[x][4] + '\n';
              var pmWeeklyNext = 'Next: ' + formatDate(pm[x][5]) + '\n';
              var pmWeeklyDays = 'Days left: ' + pm[x][6] + '\n\n';
            }
          } else if (increment == 1) {
            //monthly
            if (pm[x][3] == "Not Done") {
              var pmMonthlyLast = 'Monthly PM:\n A Monthly PM has not been done yet! \n\n';
              var pmMonthlyName = '';
              var pmMonthlyNext = '';
              var pmMonthlyDays = '';
            } else {
              var pmMonthlyLast = 'Monthly PM:\n Last: ' + formatDate(pm[x][3]);
              var pmMonthlyName = ' - ' + pm[x][4] + '\n';
              var pmMonthlyNext = 'Next: ' + formatDate(pm[x][5]) + '\n';
              var pmMonthlyDays = 'Days left: ' + pm[x][6] + '\n\n';
            }
          } else if (increment == 2) {
            //quarterly
            if (pm[x][3] == "Not Done") {
              var pm3MonthLast = '3-Month PM:\n A 3-Month PM has not been done yet!';
              var pm3MonthName = '';
              var pm3MonthNext = '';
              var pm3MonthDays = '';
            } else {
              var pm3MonthLast = '3-Month PM:\n Last: ' + formatDate(pm[x][3]);
              var pm3MonthName = ' - ' + pm[x][4] + '\n';
              var pm3MonthNext = 'Next: ' + formatDate(pm[x][5]) + '\n';
              var pm3MonthDays = 'Days left: ' + pm[x][6];
            }
          }
          increment++;
        }
      }
      
      var weeklyPMs = pmWeeklyLast + pmWeeklyName + pmWeeklyNext + pmWeeklyDays;
      var monthlyPMs = pmMonthlyLast + pmMonthlyName + pmMonthlyNext + pmMonthlyDays;
      var quarterlyPMs = pm3MonthLast + pm3MonthName + pm3MonthNext + pm3MonthDays;
      
      row = '<a href="Maintenance.html?serial=' + sn[k] + '" class="partTooltip" title="' + weeklyPMs + monthlyPMs + quarterlyPMs + '" target="_blank">' + sn[k] + '</a>';
      
    } else {
      for (var i=0, iLen=pm.length; i<iLen; i++) {
        if (pm[i][0] == sn[k]) {
          pmType = 'PM Type: ' + pm[i][2];
          pmLastRaw = pm[i][3];
          pmName = '\n Name: ' + pm[i][4];
          pmNextRaw = pm[i][5];
          pmDays = '\n Days Left: ' + pm[i][6];
          
          if (pmLastRaw != "Not Done" && pmLastRaw != "-")  {
            var pmLast = '\n Last PM: ' + formatDate(pmLastRaw);
          }
          if (pmNextRaw != "Not Done" && pmNextRaw != "-") {
            var pmNext = '\n Next PM: ' + formatDate(pmNextRaw);
          }
          row = '<a href="Maintenance.html?serial=' + sn[k] + '" class="partTooltip" title="' + pmType + pmLast + pmName + pmNext + pmDays +'" target="_blank">' + sn[k] + '</a>'
        }
      }
    }
    output.push([row]);
  }
  
  console.timeEnd("snTooltip time");
  console.log("Completed snTooltip()");
  return output;
}

// The following takes JSON data sent via POST and appends to spreadsheet
function doGet(e){
  console.log("Running doGet");
  return handleResponse(e);
}
function doPost(e){
  console.log("Running doPost");
  return handleResponse(e);
}

function handleResponse(e) {
  console.log("Running handleResponse(e): " + e);
  console.time("handleResponse time");
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.

  try {
    var sheetNameMain = loadSetting('sheetNameMain');
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameMain);
    var nextRow = mainSheet.getLastRow()+1; // get next row
    var row = [];
    
    console.log("Found last row: " + nextRow);
    
    //include timestamp
    var now = new Date();
    var date = [ now.getMonth() + 1, now.getDate(), now.getFullYear() ];
    var time = [ now.getHours(), now.getMinutes(), now.getSeconds() ];
    for ( var i = 1; i < 3; i++ ) {
      if ( time[i] < 10 ) {
        time[i] = "0" + time[i];
      }
    }
    var timestamp = date.join("/") + " " + time.join(":")
    row.push(timestamp);

    //transfer used parts and quantities into array (there is definitely an easier way to do this)
    var usedPart1 = e.parameter["usedPart1"];
    var usedPart2 = e.parameter["usedPart2"];
    var usedPart3 = e.parameter["usedPart3"];
    var usedPart4 = e.parameter["usedPart4"];
    var usedPart5 = e.parameter["usedPart5"];
    var usedPart6 = e.parameter["usedPart6"];
    var usedPart7 = e.parameter["usedPart7"];
    var usedPart8 = e.parameter["usedPart8"];
    var usedPart9 = e.parameter["usedPart9"];
    var usedPart10 = e.parameter["usedPart10"];
    var usedQty1 = e.parameter["usedQty1"];
    var usedQty2 = e.parameter["usedQty2"];
    var usedQty3 = e.parameter["usedQty3"];
    var usedQty4 = e.parameter["usedQty4"];
    var usedQty5 = e.parameter["usedQty5"];
    var usedQty6 = e.parameter["usedQty6"];
    var usedQty7 = e.parameter["usedQty7"];
    var usedQty8 = e.parameter["usedQty8"];
    var usedQty9 = e.parameter["usedQty9"];
    var usedQty10 = e.parameter["usedQty10"];
    
    var usedPartsArray = [[usedPart1,usedQty1],[usedPart2,usedQty2],[usedPart3,usedQty3],[usedPart4,usedQty4],[usedPart5,usedQty5],[usedPart6,usedQty6],[usedPart7,usedQty7],[usedPart8,usedQty8],[usedPart9,usedQty9],[usedPart10,usedQty10]];
    
    //transfer replenished parts and quantities into array
    var repPart1 = e.parameter["repPart1"];
    var repPart2 = e.parameter["repPart2"];
    var repPart3 = e.parameter["repPart3"];
    var repPart4 = e.parameter["repPart4"];
    var repPart5 = e.parameter["repPart5"];
    var repPart6 = e.parameter["repPart6"];
    var repPart7 = e.parameter["repPart7"];
    var repPart8 = e.parameter["repPart8"];
    var repPart9 = e.parameter["repPart9"];
    var repPart10 = e.parameter["repPart10"];
    var repQty1 = e.parameter["repQty1"];
    var repQty2 = e.parameter["repQty2"];
    var repQty3 = e.parameter["repQty3"];
    var repQty4 = e.parameter["repQty4"];
    var repQty5 = e.parameter["repQty5"];
    var repQty6 = e.parameter["repQty6"];
    var repQty7 = e.parameter["repQty7"];
    var repQty8 = e.parameter["repQty8"];
    var repQty9 = e.parameter["repQty9"];
    var repQty10 = e.parameter["repQty10"];
    
    var repPartsArray = [[repPart1,repQty1],[repPart2,repQty2],[repPart3,repQty3],[repPart4,repQty4],[repPart5,repQty5],[repPart6,repQty6],[repPart7,repQty7],[repPart8,repQty8],[repPart9,repQty9],[repPart10,repQty10]];
    
    //Combine used parts and qunatities into one string, separate part/qty with "x" and each entry with ", ". Also skips blank part entries
    var usedPartsFinal = "";
    for (var i = 0; i < usedPartsArray.length; i++) {
      if (usedPartsArray[i][0] != "" && usedPartsArray[i][1] != "") {
        if (usedPartsFinal == "") {
           usedPartsFinal += usedPartsArray[i][0] + "x" + usedPartsArray[i][1];
        } else {
           usedPartsFinal += ", " + usedPartsArray[i][0] + "x" + usedPartsArray[i][1];
        }
      } else {
        //do nothing 
      }
    }
	
	//Combine replenished parts and qunatities into one string, separate part/qty with "x" and each entry with ", ". Also skips blank part entries
    var repPartsFinal = "";
    for (var i = 0; i < repPartsArray.length; i++) {
      if (repPartsArray[i][0] != "" && repPartsArray[i][1] != "") {
        if (repPartsFinal == "") {
           repPartsFinal += repPartsArray[i][0] + "x" + repPartsArray[i][1];
        } else {
           repPartsFinal += ", " + repPartsArray[i][0] + "x" + repPartsArray[i][1];
        }
      } else {
        //do nothing 
      }
    }
    	
    //Should match order of columns in spreadsheet
	row.push(e.parameter["Date"]);
	row.push(e.parameter["Initials"]);
	row.push(e.parameter["SN"]);
	row.push(e.parameter["From"]);
	row.push(e.parameter["Into"]);
	row.push(e.parameter["Problem"]);
	row.push(e.parameter["Solution"]);
	row.push(usedPartsFinal);
	row.push(repPartsFinal);
	row.push(e.parameter["Notes"]);
	row.push(e.parameter["PM Type"]);
	row.push(e.parameter["FSR"]);
    
    row.push(e.parameter["SN"]);
    row.push(usedPartsFinal);
    row.push(repPartsFinal);
    
    console.log("Pushed all data into new array: " + row);
    
    //Write array to row in spreadsheet:
    mainSheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    
    console.timeEnd("handleResponse time");
    console.log("Completed handleResponse(e)");
    
    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e){
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
    var recentlySubmitted = true;
    saveSetting('recentlySubmitted',recentlySubmitted); //used to update helper link after sheet has finished calculating inventory numbers
  }
}

function monitorRecentlySubmitted() {
  //Updates last 5 rows shortly after a form submission from main web page
  //This script runs every 1 minute.
  var recentlySubmitted = loadSetting('recentlySubmitted');
  if (recentlySubmitted === 'true') {
    console.log('Form was recently submitted, updating last 5 rows');
    var sheetNameMain = loadSetting('sheetNameMain');
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameMain);
    updateHelperLinks(mainSheet.getLastRow()-4,5); //updates last 5 rows
    recentlySubmitted = false;
    saveSetting('recentlySubmitted',recentlySubmitted);
  }
}