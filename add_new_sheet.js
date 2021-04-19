/*
Sheet link: https://docs.google.com/spreadsheets/d/1wKBmqqOTsvxOGKwUU1Nvg3bbQHtdTu5abIdpyZI_WU8/edit#gid=1007564265
1. Add new blank sheet: This function is to add a new blank sheet to each client spreadsheet with a name 'Summary Sheet', and edit cell C3 in the newly added blank sheet with a formula and some format change

2. Delete sheet: This function is to delete a sheet from each client spreadsheet using sheet name. The user can delete any sheet by changing the sheet name in the function. Currently the function is set to delete sheet 'Summary Sheet'. If 'Summary Sheet' does not exist in one of the client spreadsheets, an error will occur. So the best practice is to run the 'Add new blank sheet' first and then try to play the 'Delete sheet' function.

3. Add a template: This function is to add a template sheet to each client spreadsheet. Currently the function is using 'Summary Sheet Template' in 'Product Price Test' as the template.
  Logic of this function: 
  3.1 In order to preserve the format and image in the template, here I choose to use a copy method
  3.2 Copy the template sheet from 'Product Price Test' and paste it to the destination client spreadsheet
  3.3 Move this paste sheet to the first sheet in the destination client spreadsheet. Here the user can choose the sheet index to move this sheet by changing the number in the parenthesis in this line of code 'destSheet.moveActiveSheet(1);'
  3.4 Since this paste sheet is a copy of the template, so the sheet name needed to be changed. So the last step is to rename this sheet to 'Summary Sheet'

4. Add a row in template: This function is to add a new row at certain row index in 'Summary Sheet' in each client spreadsheet. Currently the function is set to add a row at index 13 and set cell(B13) to a red colored text input. The user can choose any row index to add a new row by changing the number in the parenthesis in this line of code 'destSheet.getSheetByName("Summary Sheet").insertRows(13);' and can also choose which cell to edit by changing the row index and column index in this line of code 'destSheet.getSheetByName("Summary Sheet").getRange(13,2).setValue("You can input anything here").setFontColor('red');'. If the 'Summary Sheet' does not exist in one of the client spreadsheets, an error will occur. So the best practice is to run the 'Add new blank sheet' or 'Add a template' function first and then try to play the 'Add a row' function.

*/



//Creat menu for export
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Export')
      .addItem('Add new blank sheet', 'addNewSheet')
      .addItem('Delete sheet','DeleteSheet')
      .addItem('Add a template','AddTemplate')
      .addItem('Add a row and edit cell','AddRows')
      .addToUi();
}


function addNewSheet(){
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet().getSheetByName("Client Balance");
  var totalrow = sheet.getLastRow();
  for ( var r = 2; r <= totalrow; r++){
    var url = sheet.getRange(r,2).getValue();
    var clientSheet = app.openByUrl(url);
    //Add new blank sheet
    clientSheet.insertSheet('Summary Sheet');
    //Add formula in the destination cell in 'Summary Sheet'. getRange(row,col)is the cell location, getSheetByName() is the sheet location
    var cell = clientSheet.getSheetByName('Summary Sheet').getRange(3,3).setFormula("=SUM(3+4)");
    //Change cell format
    cell.setBackground("yellow");
    cell.setFontColor('red');

  }
}

function DeleteSheet(){
  var app = SpreadsheetApp;
  var sheet = app.getActiveSpreadsheet().getSheetByName("Client Balance");
  var totalrow = sheet.getLastRow();
  for ( var r = 2; r <= totalrow; r++){
    var url = sheet.getRange(r,2).getValue();
    var clientSheet = app.openByUrl(url);
    //Delete sheet 'Summary Sheet'. The user can delete anysheet by changing the sheetname in the parenthesis 
    var dsheet = clientSheet.getSheetByName("Summary Sheet");
    clientSheet.deleteSheet(dsheet);
  }
}


function AddTemplate(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Client Balance");
  var templateSheet = ss.getSheetByName('Summary Sheet'); //Get the template sheet
  var totalrow = sheet.getLastRow();
  for (var r = 2; r <= totalrow; r++){
    var url = sheet.getRange(r,2).getValue();
    var destSheet = SpreadsheetApp.openByUrl(url);
    templateSheet.copyTo(destSheet); //Paste the template sheet to destination spreadsheet

    destSheet.getSheetByName("Copy of Summary Sheet").activate();
    destSheet.moveActiveSheet(1); //Move the pasted template sheet to sheet at Index 
    destSheet.getSheetByName("Copy of Summary Sheet").setName("Summary Sheet"); //Rename the sheet 
  }
}


function AddRows(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Client Balance");
  var totalrow = sheet.getLastRow();
  for (var r = 2; r <= totalrow; r++){
    var url = sheet.getRange(r,2).getValue();
    var destSheet = SpreadsheetApp.openByUrl(url);
    //destSheet.getSheetByName("Product prices").insertRows(13); //Insert a row at Row Index 13
    destSheet.getSheetByName("Product prices").getRange(1,22).setValue("Czechia Yunexpress").setFontColor('black'); //Edit a new cell in the the get range cell
  
  }
}



