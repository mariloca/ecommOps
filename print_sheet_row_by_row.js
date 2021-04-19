//Sheet link: https://docs.google.com/spreadsheets/d/1trJ-2F-dkZaZes8EUBns2ift4JzB2NrEwIatAt6_0fI/edit#gid=268970711


//Print sheet row by row
function saveAsCSVinRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName('Quote Database');

  /*
  var folderecommos = DriveApp.getFolderById("17EQy-rErLDwf7vJVrhTkgNOS0X6yrsZC");
  // convert all available sheet data to csv format
  var resFromCSV = convertRangeToCsvFile_inRow(sheet);
  for (var key in resFromCSV){
    var filename = key + ".csv";
    var file = folderecommos.createFile(filename,resFromCSV[key]);
  }
 */

  // convert all available sheet data to csv format
  var resFromCSV = convertRangeToCsvFile_inRow(sheet);
  for (var key in resFromCSV){
    var folder = DriveApp.createFolder(key);
    var filename = key + ".csv";
    var file = folder.createFile(filename,resFromCSV[key]);
  }


  /*
  var folder = DriveApp.createFolder(foldername);
  // create a file in the Docs List with the given name and the csv data
  var file = folder.createFile(fileName, csvFile);
  */
}


function convertRangeToCsvFile_inRow(sheet) {
  var csvList = {};
  var totalcolumn = sheet.getLastColumn();
  var totalrow = sheet.getLastRow();

  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var template = ss.getSheetByName('Print Template');

  var tmpfileName = sheet.getRange(4,2).getValue();
  clearTemplate();
  for (var r = 4; r<=totalrow;r++)
  {
    if (sheet.getRange(r,1).getValue() == "Yes")
    {
      if (sheet.getRange(r,2).getValue() == tmpfileName)
      {
        var copyRange = sheet.getRange(r,1,1,totalcolumn);
        copyRange.copyTo(template.getRange(template.getLastRow()+1,1),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
      }
      else {
        var fileName = tmpfileName;
        tmpfileName = sheet.getRange(r,2).getValue();
        var startRowIndex = r;
        template.getRange("C:C").setNumberFormat('M/d/yyyy');
        var activeRange = template.getDataRange();
        try{
          var data = activeRange.getValues();
          var csvFile = undefined;

          // loop through the data in the range and build a string with the csv data
          if (data.length > 1) {
            var csv = "";
            for (var row = 0; row < data.length; row++) {
              for (var col = 0; col < data[row].length; col++) {
                if (data[row][col].toString().indexOf(",") != -1) {
                  data[row][col] = "\"" + data[row][col] + "\"";
                }
              }

              // join each row's columns
              // add a carriage return to end of each row, except for the last one
              if (row < data.length-1) {
                csv += data[row].join(",") + "\r\n";
              }
              else {
                csv += data[row];
              }
            }
            csvFile = csv;
          }
          csvList[fileName] = csvFile;
          clearTemplate();
          copyRange = sheet.getRange(r,1,1,totalcolumn);
          copyRange.copyTo(template.getRange(template.getLastRow()+1,1),SpreadsheetApp.CopyPasteType.PASTE_VALUES,false);
        }
        catch(err){
          Logger.log(err);
          Browser.msgBox(err);
        }    
      }      
    }  
  }
  //Last print out
  var fileName = tmpfileName;
  template.getRange("C:C").setNumberFormat('M/d/yyyy');
  var activeRange = template.getDataRange();
  try{  
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }
        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    csvList[fileName] = csvFile;
    clearTemplate();
  }
  catch(err){
    Logger.log(err);
    Browser.msgBox(err);
  } 
  
  return csvList;
}


function clearTemplate(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var template = ss.getSheetByName('Print Template');
  template.clearContents();
  //add header
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('1:2').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Print Template'), true);
  spreadsheet.getRange('\'Quote Database\'!1:2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}


//========================================================
function printQuote(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var ActiveSheetName = ss.getActiveSheet().getName();
  if (ActiveSheetName == "Quote Template"){
    var PrintSheet = ss.getSheetByName("Quote Template");
  } else if (ActiveSheetName == "Multi Unit Quote Template"){
    var PrintSheet = ss.getSheetByName("Multi Unit Quote Template");
  } else {
    Logger.log("Not the right quote sheet.");
    Browser.msgBox("Not the right quote sheet. Please select the right quote sheet");
  }

  //save to ecommos folder
  var folderecommos = DriveApp.getFolderById("17EQy-rErLDwf7vJVrhTkgNOS0X6yrsZC");
  // convert all available sheet data to csv format
  var csvFile = convertToCSV(PrintSheet);
  var filename = "Normal Quote.csv";
  //var filename = "Multi unit Quote.csv";
  var file = folderecommos.createFile(filename,csvFile);

  /* save to new folder
  var folder = DriveApp.createFolder(foldername);
  // create a file in the Docs List with the given name and the csv data
  var file = folder.createFile(fileName, csvFile);
  // convert all available sheet data to csv format
  var resFromCSV = convertToCSV(sheet);
  for (var key in resFromCSV){
    var folder = DriveApp.createFolder(key);
    var filename = key + ".csv";
    var file = folder.createFile(filename,resFromCSV[key]);
  }
  */
}

function convertToCSV(sheet) {
  var activeRange = sheet.getRange(1,1,75,15);
  //var activeRange = sheet.getDataRange();
  try{
    
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
      }
  }
  catch(err){
    Logger.log(err);
    Browser.msgBox(err);
  }
  return csvFile;
}










  