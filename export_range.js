//Sheet link: https://docs.google.com/spreadsheets/d/1Z1WU8w0D7gWhT4JhpNs9B3dNFI8xI61i4S8lRTMhZ0k/edit#gid=119345874

//Export selected data range from a sheet to a different spreadsheet
function exportRange(){
  var app = SpreadsheetApp;
  var exportFromSheet = app.getActiveSpreadsheet().getSheetByName("Client quotes");
  var exportToSheet = app.openByUrl(
      'https://docs.google.com/spreadsheets/d/1wKBmqqOTsvxOGKwUU1Nvg3bbQHtdTu5abIdpyZI_WU8/edit#gid=1969876936').getSheetByName("Sheet4");
  
  //var exportToSheet = app.getActiveSpreadsheet().getSheetByName("Test");
  
  var quoteTotalRow = exportFromSheet.getLastRow();
  var quoteTotalColumn = exportFromSheet.getLastColumn();
  var colIndex = 0
  
  //Find starting column index:colIndex
  for (var c = 1; c <= quoteTotalColumn; c++){
    if (exportFromSheet.getRange(1,c).getValue() === "Name"){
        colIndex = c;
        break;
      }
    }
    
  //Find copy range and save values in a 2D array
  var copyRange = exportFromSheet.getRange(3, colIndex, quoteTotalRow-2,quoteTotalColumn-1);
  var copyValues = copyRange.getValues();
  
  //Write the 2D array range to destination range
  var targetRange = exportToSheet.getRange(exportToSheet.getLastRow()+1, 1,quoteTotalRow-2,quoteTotalColumn-1);
  targetRange.setValues(copyValues);

  /*When copy to same spreadsheet, can use "copyTo" method
  copyRange.copyTo(exportToSheet.getRange(exportToSheet.getLastRow()+1, 1)); 
  */
}



//Export selected data range to from a sheet to another sheet within the same spreadsheet
function exportRangetoSameSpreadsheet(){
  var app = SpreadsheetApp;
  var exportFromSheet = app.getActiveSpreadsheet().getSheetByName("Client quotes");
  
  var exportToSheet = app.getActiveSpreadsheet().getSheetByName("Test");
  
  var quoteTotalRow = exportFromSheet.getLastRow();
  var quoteTotalColumn = exportFromSheet.getLastColumn();
  var colIndex = 0
  
  //Find starting column index:colIndex
  for (var c = 1; c <= quoteTotalColumn; c++){
    if (exportFromSheet.getRange(1,c).getValue() === "Name"){
        colIndex = c;
        break;
      }
    }
    
  //Find copy range 
  var copyRange = exportFromSheet.getRange(3, colIndex, quoteTotalRow-2,quoteTotalColumn-1);  
  //When copy to same spreadsheet, can use "copyTo" method
  copyRange.copyTo(exportToSheet.getRange(exportToSheet.getLastRow()+1, 1)); 
 
}



