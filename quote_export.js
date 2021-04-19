/*
Sheet link: https://docs.google.com/spreadsheets/d/1trJ-2F-dkZaZes8EUBns2ift4JzB2NrEwIatAt6_0fI/edit#gid=268970711
QuoteExport: export quote from "Normal Quote Sheet" or "Multi Unit Quote Sheet" to "Quote Database"
1. In order to run the 'QuoteExport' function, the user has to be at either the "Normal Quote Sheet" or the "Multi Unit Quote Sheet"; otherwise, browser will pop out alert window to remind
  the user to select the right quote sheet;
2. Once the user is at one of the two quote sheets, the user can click the 'Quote' tab from the menu bar to export the quote to database;
3. After the script finished running, the export process is then completed;
4. Be sure to clear the already exported quote in the sheet before next export to prevent export repetitive quote data into the database.
5. Logic of 'QuoteExport':
  5.1 For Normal quote: the script will export quote from Row4 to LastRow in "Normal Quote Sheet" to "Quote Database" starting with LastRow in "Quote Database";
  5.2 For Multi quote: the script will export quote from Row6 to LastRow in "Multi Unit Quote Sheet" to "Quote Database" starting with LastRow in "Quote Database";


PrintQuote: print quote in "Quote Template" or "Multi Unit Quote Template" to the user's Google Drive 
1. In order to run the 'PrintQuote' function, the user has to be at either the "Quote Template" or the "Multi Unit Quote Template"; otherwise, browser will pop out alert window to remind
  the user to select the right template sheet;
2. Logic of 'PrintQuote':
  2.1 For Normal quote: Once the function finished running, the user's Google Drive will have a new spreadsheet contains Normal quote data with the name stored in Cell(B3) in "Normal Quote Sheet"
  2.2 For Multi quote: Once the function finished running, the user's Google Drive will have a new spreadsheet contains Multi unit quote data with the name stored in Cell(B3) in "Multi Unit Quote Sheet"
*/

//Creat menu for quote
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Quote')
      .addItem('Export Quote to Quote Database', 'QuoteExport')
      .addItem('Print Quote', 'PrintQuote')
      .addToUi()
}

//Runtime of one row with Regular,Upsell,Multi: 80s
function QuoteExport(){ 
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet(); 
  var ActiveSheetName = ss.getActiveSheet().getName();
  if (ActiveSheetName == "Normal Quote Sheet"){
    var exportFromSheet = ss.getSheetByName("Normal Quote Sheet");
  } else if (ActiveSheetName == "Multi Unit Quote Sheet"){
      exportFromSheet = ss.getSheetByName("Multi Unit Quote Sheet");
  } else {
    Browser.msgBox("Not the right quote sheet. Please select the right quote sheet");
    return;
  }

  var lastrow = exportFromSheet.getLastRow();
  var lastcol = exportFromSheet.getLastColumn();
  var exportToSheet = app.getActiveSpreadsheet().getSheetByName("Quote Database");
  var totalcolumn = exportToSheet.getLastColumn();
  var totalrow = exportToSheet.getLastRow();
  
  for (var c = 1; c <= lastcol; c++){
    if (exportFromSheet.getRange(1,c).getValue() =="Date 日期"){
      var DateIndex = c;
    } else if (exportFromSheet.getRange(1,c).getValue() == "English Sku name"){
      var SkuNameIndex = c;
    }
    else if (exportFromSheet.getRange(1,c).getValue() == "Notes"){
      var NotesIndex = c;
    }
    else if (exportFromSheet.getRange(1,c).getValue() == "Multi quote?"){
      var MultiunitsIndex = c;  
    }
    else if (exportFromSheet.getRange(1,c).getValue() == "Yun Vol Reg?"){
      var YunVolRegIndex = c; 
    }
  }
  
  //Store headers' index in "Quote sheet" in a hashtable for future lookup
  var headerMap = {}; 
  for (var col = 1; col <= lastcol; col++){
    var header = exportFromSheet.getRange(1,col).getValue();
    headerMap[header] = col;
    }   
  var ToheaderMap = {};
  for (var i = 1; i <= totalcolumn; i++){
    var header = exportToSheet.getRange(1,i).getValue();
    ToheaderMap[header]=i;
  }

  //Initialize row index for export
  if (app.getActiveSheet().getSheetName()=="Normal Quote Sheet"){
    var row = 4;
  } else if (app.getActiveSheet().getSheetName()=="Multi Unit Quote Sheet"){
    row = 6;
  }

  //FIRST HALF EXPORT: DATES ~ NOTES
  var OutputRowIndex = exportToSheet.getLastRow()+1;
  var exportRowNum = lastrow - row+1;
  var FirstexportColumnsNum = NotesIndex-DateIndex+1;
  var FirstcopyRange = exportFromSheet.getRange(row, 1, exportRowNum, FirstexportColumnsNum);  
  FirstcopyRange.copyTo(exportToSheet.getRange(OutputRowIndex, 3)); 

  //SECOND HALF EXPORT : MULTI QUOTE? ~ YUN VOL REG  
  var SecexportColumnsNum = YunVolRegIndex-MultiunitsIndex+1;
  var SeccopyRange = exportFromSheet.getRange(row, MultiunitsIndex, exportRowNum, SecexportColumnsNum); 
  SeccopyRange.copyTo(exportToSheet.getRange(OutputRowIndex, ToheaderMap["Multi quote?"]));

  //LAST COLUMN EXPORT:'English Sku Name'
  var SkuCopyRange = exportFromSheet.getRange(row,SkuNameIndex,exportRowNum,1);
  SkuCopyRange.copyTo(exportToSheet.getRange(OutputRowIndex,ToheaderMap["English Sku name"]));
  
}
    


function PrintQuote() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var ActiveSheetName = ss.getActiveSheet().getName();
  if (ActiveSheetName == "Quote Template"){
    var sourceSheet = ss.getSheetByName("Quote Template");
    var destSheetName = ss.getSheetByName("Normal Quote Sheet").getRange(3,2).getValue(); //Get the output spreadsheet name at this cell 
  } else if (ActiveSheetName == "Multi Unit Quote Template"){
      sourceSheet = ss.getSheetByName("Multi Unit Quote Template");
      destSheetName = ss.getSheetByName("Multi Unit Quote Sheet").getRange(3,2).getValue(); //Get the output spreadsheet name at this cell 
  } else {
    Browser.msgBox("Not the right template sheet. Please select the right template sheet");
    return;
  }

  //Only create a spreadsheet with 'destSheetName' and store the spreadsheet directly to the user's Google Drive
  var destSpreadsheet = SpreadsheetApp.open(DriveApp.getFileById(ss.getId()).makeCopy(destSheetName)); //Copy whole spreadsheet

  //Delete redundant sheets other than the Quote template sheet
  var sheets = destSpreadsheet.getSheets();
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() != sourceSheet.getName()){
    destSpreadsheet.deleteSheet(sheets[i]);
    }
  }
  destSpreadsheet.getSheets()[0].setName("Quote");
}
