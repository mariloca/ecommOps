/*
Sheet link: https://docs.google.com/spreadsheets/d/1trJ-2F-dkZaZes8EUBns2ift4JzB2NrEwIatAt6_0fI/edit#gid=268970711

Price Quoting Category:
  1. Regular (default, 1 unit quote)
  2. Upsell (Basically the default minus the base price, the price that we would charge if a product i sent with another product in the same   package, so they can "share" the base price and therefore be cheaper)
  3. Multi (1 to 10 unit quotes)

Requirement for RegularQuote/UpsellQuote/MultiQuote:
  1. Column headers in "Quote Sheet" must be the same in the "Quote Database" sheet
  2. The order of Columns A~J in "Quote Sheet" must be the same in the "Quote Database" sheet

Quoting:
  1. The original quote data is always from the last row of "Quote Sheet"
*/


//Runtime of one row with Regular,Upsell,Multi: 80s
function QuoteByBunch(){ 
  var app = SpreadsheetApp;
  var exportFromSheet = app.getActiveSpreadsheet().getSheetByName("Quote Sheet");
  var lastrow = exportFromSheet.getLastRow();
  var lastcol = exportFromSheet.getLastColumn();
  var exportToSheet = app.getActiveSpreadsheet().getSheetByName("Quote Database");
  var totalcolumn = exportToSheet.getLastColumn();
  var totalrow = exportToSheet.getLastRow();
  
  for (var c = 1; c <= totalcolumn; c++){
    if (exportToSheet.getRange(1,c).getValue() =="Date 日期"){
      var DateIndex = c;
    } 
    else if (exportToSheet.getRange(1,c).getValue() =="Margin per unit"){
      var MarginPerUnitIndex = c;
    }
    else if (exportToSheet.getRange(1,c).getValue() == "English Sku name"){
      var exportEndColIndex = c;
    }
    else if (exportToSheet.getRange(1,c).getValue() == "Notes"){
      var NotesIndex = c;
    }
    else if (exportToSheet.getRange(1,c).getValue() == "Multi quote units"){
      var MultiunitsIndex = c;  
    }
    else if (exportToSheet.getRange(1,c).getValue() == "中文名字"){
      var ChinesenameIndex = c; 
    }
    else if (exportToSheet.getRange(1,c).getValue() == "最低利润 min."){
      var MinIndex = c; 
    }
  }
  //store headers' index in "Quote sheet" in a hashtable for future lookup
  var headerMap = {}; 
  for (var col = 1; col <= lastcol; col++){
    var header = exportFromSheet.getRange(1,col).getValue();
    headerMap[header] = col;
    }   
  var ToheaderMap = {};
  for (var i = NotesIndex; i <= exportEndColIndex; i++){
    var header = exportToSheet.getRange(1,i).getValue();
    ToheaderMap[header]=i;
  }

  for(var row=2;row <= lastrow;row++){
    //FIRST HALF EXPORT: DATES ~ MARGIN PER UNIT
    var RegularOutputRow = exportToSheet.getLastRow()+1;
    var exportColumnsNum = MarginPerUnitIndex-DateIndex+1;
    var copyRange = exportFromSheet.getRange(row, 1, 1, exportColumnsNum);  
    copyRange.copyTo(exportToSheet.getRange(RegularOutputRow, DateIndex)); 

    //SECOND HALF EXPORT : NOTES ~ ENGLISH SKU NAME   
    //Find left empty col index of TOSheet in FROMsheet and setvalue to its correspondent cell
    for (var i = NotesIndex; i<= exportEndColIndex; i++){
      var Toheader = exportToSheet.getRange(1,i).getValue();                                //get header in TOSheet
      if ( Toheader in headerMap == true){
        var setcell = exportFromSheet.getRange(row,headerMap[Toheader]).getValue();     //get cell value in FROMsheet
        exportToSheet.getRange(RegularOutputRow, ToheaderMap[Toheader]).setValue(setcell);
      }
    } 

    //UPSELL
    if ((exportFromSheet.getRange(row,headerMap["Upsell quote?"]).getValue()) == "Yes"){
      var UpsellOutputRow = exportToSheet.getLastRow()+1;
      //FIRST HALF EXPORT: DATES ~ MARGIN PER UNIT
      var copyRange = exportFromSheet.getRange(row, 1, 1, exportColumnsNum);  
      copyRange.copyTo(exportToSheet.getRange(UpsellOutputRow, DateIndex)); 
      
      //SECOND HALF EXPORT : NOTES ~ ENGLISH SKU NAME
      //NOTES ~ Sku name-1. Find left empty col index of TOSheet in FROMsheet and setvalue to its correspondent cell
      for (var i = NotesIndex; i<= exportEndColIndex-1; i++){
        var Toheader = exportToSheet.getRange(1,i).getValue();                                //get header in TOSheet
        if (Toheader in headerMap == true){
            var setcell = exportFromSheet.getRange(row,headerMap[Toheader]).getValue();     //get cell value in FROMsheet
            exportToSheet.getRange(UpsellOutputRow, ToheaderMap[Toheader]).setValue(setcell);        
        }
      } 
      //Set 'English SKU Name' column to name+upsell
      var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+" "+"Upsell";
      exportToSheet.getRange(UpsellOutputRow, exportEndColIndex).setValue(exportcell);
    }

    //MULTIQUOTE
    if ((exportFromSheet.getRange(row,headerMap["Multi quote?"]).getValue()) == "Yes"){
      var MultiOutputRow = exportToSheet.getLastRow()+1;
      //FIRST HALF EXPORT: DATES ~ MARGIN PER UNIT
      var copyRange = exportFromSheet.getRange(row, 1, 1, exportColumnsNum);  
      copyRange.copyTo(exportToSheet.getRange(MultiOutputRow, DateIndex)); 
      
      //SECOND HALF EXPORT : NOTES ~ ENGLISH SKU NAME
      //NOTES ~ Sku name-1. Find left empty col index of TOSheet in FROMsheet and setvalue to its correspondent cell
      for (var i = NotesIndex; i<= exportEndColIndex-1; i++){
        var Toheader = exportToSheet.getRange(1,i).getValue();                                //get header in TOSheet
        if (Toheader in headerMap == true){
            var setcell = exportFromSheet.getRange(row,headerMap[Toheader]).getValue();     //get cell value in FROMsheet
            exportToSheet.getRange(MultiOutputRow, ToheaderMap[Toheader]).setValue(setcell);     
        }
      } 
      //FIRST ROW OF MULTI QUOTE 'English SKU Name' column to name-1
      var units = 1
      exportToSheet.getRange(MultiOutputRow,MultiunitsIndex).setValue(units);
      var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+"-"+ units;
      exportToSheet.getRange(MultiOutputRow, exportEndColIndex).setValue(exportcell);

      var marginperunit = exportToSheet.getRange(exportToSheet.getLastRow(),MarginPerUnitIndex).getValue();
      //2~10 ROWS OF MULTI QUOTE
      for (var i=1; i<=9; i++){
        //First 4 columns (DATES, CLIENT, LINK, PRODUCT CHINESE NAME)
        var exportColumnsNum = ChinesenameIndex-DateIndex+1;
        var copyRange = exportFromSheet.getRange(row, 1, 1, exportColumnsNum);  
        copyRange.copyTo(exportToSheet.getRange(exportToSheet.getLastRow()+1, DateIndex));

        //Second part -calculated price (COG, WEIGHT,MIN,GOAL,ACTUAL MARGIN)
        for (var k = ChinesenameIndex+1; k<=ChinesenameIndex+2; k++){
          var baseInfo = exportToSheet.getRange(exportToSheet.getLastRow()-i,k).getValue();
          exportToSheet.getRange(exportToSheet.getLastRow(),k).setValue(baseInfo+baseInfo*i);
        }
        for (var m = MinIndex; m<=MarginPerUnitIndex-1; m++){
          exportToSheet.getRange(exportToSheet.getLastRow(),m).setValue(exportToSheet.getRange(exportToSheet.getLastRow()-i,m).getValue()+marginperunit*i);
        }
        //Margin Per Unit remain same
        exportToSheet.getRange(exportToSheet.getLastRow(),MarginPerUnitIndex).setValue(marginperunit);

        //Third Part -Units
        exportToSheet.getRange(exportToSheet.getLastRow(),MarginPerUnitIndex+1).setValue(i+1);

        //Forth Part - static copy paste (NOTES, YUN BASIC, YUN VOL REG)
        for (var n = NotesIndex; n<= exportEndColIndex-1; n++){
          var Toheader = exportToSheet.getRange(1,n).getValue();                                
          if (Toheader in headerMap == true){
            var setcell = exportFromSheet.getRange(row,headerMap[Toheader]).getValue();     
            exportToSheet.getRange(exportToSheet.getLastRow(), ToheaderMap[Toheader]).setValue(setcell);
          }
        }
        //Last Part - update Sku name with new units (ENGLISH SKU NAME)
        var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+"-"+ (i+1);
        exportToSheet.getRange(exportToSheet.getLastRow(), exportEndColIndex).setValue(exportcell);
      } 
    }
  }
}





//Runtime of one row with Regular,Upsell,Multi: 138s
function QuoteByColumn(){ 
  var app = SpreadsheetApp;
  var exportFromSheet = app.getActiveSpreadsheet().getSheetByName("Quote Sheet");
  var lastrow = exportFromSheet.getLastRow();
  var lastcol = exportFromSheet.getLastColumn();
  var exportToSheet = app.getActiveSpreadsheet().getSheetByName("Quote Database");
  var totalcolumn = exportToSheet.getLastColumn();
  var totalrow = exportToSheet.getLastRow();

  for (var c = 1; c <= totalcolumn; c++){
    if (exportToSheet.getRange(1,c).getValue() == "English Sku name"){
      var exportEndColIndex = c;
    }
  }

  //store headers' index in "Quote sheet" in a hashtable for future lookup
  var ToheaderMap={};
  for (var m = 1; m<=totalcolumn;m++){
    var toheader = exportToSheet.getRange(1,m).getValue();
    ToheaderMap[toheader]=m;
  }
  //store headers' index in "Quote sheet" in a hashtable for future lookup
  var headerMap = {}; 
  for (var col = 1; col <= lastcol; col++){
    var header = exportFromSheet.getRange(1,col).getValue();
    headerMap[header] = col;
    }

  for (var row = 2; row <= lastrow; row++){
    //REGULAR
    //Loop columns in TOSheet to check if exist in 'headerMap' in FROMSheet then export cell by cell
    var RegularOutputRow = exportToSheet.getLastRow()+1;
    for (var k = 1; k <= totalcolumn; k++){
      if (headerMap[exportToSheet.getRange(1,k).getValue()]!=null){
        var cell = exportFromSheet.getRange(row,headerMap[exportToSheet.getRange(1,k).getValue()]).getValue();
        exportToSheet.getRange(RegularOutputRow,k).setValue(cell);
      }
    }

    //UPSELL
    if ((exportFromSheet.getRange(row,headerMap["Upsell quote?"]).getValue()) == "Yes"){
      var UpsellOutputRow = exportToSheet.getLastRow()+1;
      //First half
      for (var k = 1; k <= totalcolumn; k++){
        if (headerMap[exportToSheet.getRange(1,k).getValue()]!=null){
          var cell = exportFromSheet.getRange(row,headerMap[exportToSheet.getRange(1,k).getValue()]).getValue();
          exportToSheet.getRange(UpsellOutputRow,k).setValue(cell);
        }
      }
      //Set 'English SKU Name' column to name+upsell 
      var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+" "+"Upsell";
      exportToSheet.getRange(UpsellOutputRow,exportEndColIndex).setValue(exportcell);
    }

    //MULTISELL
    if ((exportFromSheet.getRange(row,headerMap["Multi quote?"]).getValue()) == "Yes"){
      var MultiOutputRow = exportToSheet.getLastRow()+1;
      //Loop columns in TOSheet to check if exist in 'headerMap' in FROMSheet and export cell by cell from 'Date'~('SKU NAME'-1)
      for (var k = 1; k <= totalcolumn; k++){
        if (headerMap[exportToSheet.getRange(1,k).getValue()]!=null){
          var cell = exportFromSheet.getRange(row,headerMap[exportToSheet.getRange(1,k).getValue()]).getValue();
          exportToSheet.getRange(MultiOutputRow,k).setValue(cell);
        }
      }
      //FIRST ROW OF MULTI QUOTE 'English SKU Name' column to name-1
      var units = 1
      exportToSheet.getRange(MultiOutputRow,ToheaderMap["Multi quote units"]).setValue(units);
      var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+"-"+ units;
      exportToSheet.getRange(MultiOutputRow, ToheaderMap["English Sku name"]).setValue(exportcell);
  
      //SECOND ROW ~ 10th ROW
      var marginperunit = exportToSheet.getRange(exportToSheet.getLastRow(),ToheaderMap["Margin per unit"]).getValue();
      for (var i = 1; i<= 9; i++){
        var currentrow = exportToSheet.getLastRow()+1;
        for (var col = 1; col <= totalcolumn; col++){
          var colheader = exportToSheet.getRange(1,col).getValue();
          if (headerMap[colheader]!=null){
            if (colheader == "产品成本 COG"){
              var basecog = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(basecog+basecog*i);
            }
            else if (colheader == "重量 Weight"){
              var baseweight = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(baseweight+baseweight*i);         
            }
            else if (colheader =="最低利润 min."){
              var basemin = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(basemin+marginperunit*i);
            }
            else if (colheader == "理想利润 Goal"){
              var basegoal = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(basegoal+marginperunit*i);
            }
            else if (colheader == "Actual Margin"){
              var basemargin = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(basemargin+marginperunit*i);
            }
            else if (colheader == "Margin per unit"){
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(marginperunit);
            }
            else if (colheader == "English Sku name"){
              var exportcell = exportFromSheet.getRange(row,headerMap[colheader]).getValue()+"-"+ (i+1);
              exportToSheet.getRange(currentrow, ToheaderMap[colheader]).setValue(exportcell);          
            }
            else {
              var cell = exportFromSheet.getRange(row,headerMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,col).setValue(cell);
            }
          }
        }
        //Column 'Multi quote units' adds 1
        exportToSheet.getRange(currentrow,ToheaderMap["Multi quote units"]).setValue(i+1);
      }
    }
  }
}


//============Print row by row=============================================================
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




//===============================================================
//4 secs
function QuoteByBunch(){ 
  var app = SpreadsheetApp;
  var exportFromSheet = app.getActiveSpreadsheet().getSheetByName("Quote Sheet");
  var lastrow = exportFromSheet.getLastRow();
  var lastcol = exportFromSheet.getLastColumn();
  var exportToSheet = app.getActiveSpreadsheet().getSheetByName("Quote Database");
  var totalcolumn = exportToSheet.getLastColumn();
  var totalrow = exportToSheet.getLastRow();
  
  for (var c = 1; c <= totalcolumn; c++){
    if (exportToSheet.getRange(1,c).getValue() =="Date 日期"){
      var DateIndex = c;
    } 
    else if (exportToSheet.getRange(1,c).getValue() =="Margin per unit"){
      var MarginPerUnitIndex = c;
    }
    else if (exportToSheet.getRange(1,c).getValue() == "English Sku name"){
      var exportEndColIndex = c;
    }
    else if (exportToSheet.getRange(1,c).getValue() == "Notes"){
      var NotesIndex = c;
    }
    else if (exportToSheet.getRange(1,c).getValue() == "Multi quote units"){
      var MultiunitsIndex = c;  
    }
    else if (exportToSheet.getRange(1,c).getValue() == "中文名字"){
      var ChinesenameIndex = c; 
    }
    else if (exportToSheet.getRange(1,c).getValue() == "最低利润 min."){
      var MinIndex = c; 
    }
  }
  //store headers' index in "Quote sheet" in a hashtable for future lookup
  var headerMap = {}; 
  for (var col = 1; col <= lastcol; col++){
    var header = exportFromSheet.getRange(1,col).getValue();
    headerMap[header] = col;
    }
    
  var ToheaderMap = {};
  for (var i = NotesIndex; i <= exportEndColIndex; i++){
    var header = exportToSheet.getRange(1,i).getValue();
    ToheaderMap[header]=i;
  }


  for(var row=2;row <= lastrow;row++){
    //FIRST HALF EXPORT: DATES ~ MARGIN PER UNIT
    var RegularOutputRow = exportToSheet.getLastRow()+1;
    var exportColumnsNum = MarginPerUnitIndex-DateIndex+1;
    var copyRange = exportFromSheet.getRange(row, 1, 1, exportColumnsNum);  
    copyRange.copyTo(exportToSheet.getRange(RegularOutputRow, DateIndex)); 

    //SECOND HALF EXPORT : NOTES ~ ENGLISH SKU NAME
    
    //Find left empty col index of TOSheet in FROMsheet and setvalue to its correspondent cell
    for (var i = NotesIndex; i<= exportEndColIndex; i++){
      var Toheader = exportToSheet.getRange(1,i).getValue();                                //get header in TOSheet
      if ( Toheader in headerMap == true){
        var setcell = exportFromSheet.getRange(row,headerMap[Toheader]).getValue();     //get cell value in FROMsheet
        exportToSheet.getRange(RegularOutputRow, ToheaderMap[Toheader]).setValue(setcell);
      }
    } 

    //UPSELL
    if ((exportFromSheet.getRange(row,headerMap["Upsell quote?"]).getValue()) == "Yes"){
      var UpsellOutputRow = exportToSheet.getLastRow()+1;
      //FIRST HALF EXPORT: DATES ~ MARGIN PER UNIT
      var copyRange = exportFromSheet.getRange(row, 1, 1, exportColumnsNum);  
      copyRange.copyTo(exportToSheet.getRange(UpsellOutputRow, DateIndex)); 
      
      //SECOND HALF EXPORT : NOTES ~ ENGLISH SKU NAME
      //NOTES ~ Sku name-1. Find left empty col index of TOSheet in FROMsheet and setvalue to its correspondent cell
      for (var i = NotesIndex; i<= exportEndColIndex-1; i++){
        var Toheader = exportToSheet.getRange(1,i).getValue();                                //get header in TOSheet
        if (Toheader in headerMap == true){
            var setcell = exportFromSheet.getRange(row,headerMap[Toheader]).getValue();     //get cell value in FROMsheet
            exportToSheet.getRange(UpsellOutputRow, ToheaderMap[Toheader]).setValue(setcell);        
        }
      } 
      //Set 'English SKU Name' column to name+upsell
      var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+" "+"Upsell";
      exportToSheet.getRange(UpsellOutputRow, exportEndColIndex).setValue(exportcell);
    }

    //MULTIQUOTE
    if ((exportFromSheet.getRange(row,headerMap["Multi quote?"]).getValue()) == "Yes"){
      var MultiOutputRow = exportToSheet.getLastRow()+1;
      //FIRST HALF EXPORT: DATES ~ MARGIN PER UNIT
      var copyRange = exportFromSheet.getRange(row, 1, 1, exportColumnsNum);  
      copyRange.copyTo(exportToSheet.getRange(MultiOutputRow, DateIndex)); 
      
      //SECOND HALF EXPORT : NOTES ~ ENGLISH SKU NAME
      //NOTES ~ Sku name-1. Find left empty col index of TOSheet in FROMsheet and setvalue to its correspondent cell
      for (var i = NotesIndex; i<= exportEndColIndex-1; i++){
        var Toheader = exportToSheet.getRange(1,i).getValue();                                //get header in TOSheet
        if (Toheader in headerMap == true){
            var setcell = exportFromSheet.getRange(row,headerMap[Toheader]).getValue();     //get cell value in FROMsheet
            exportToSheet.getRange(MultiOutputRow, ToheaderMap[Toheader]).setValue(setcell);     
        }
      } 
      //FIRST ROW OF MULTI QUOTE 'English SKU Name' column to name-1
      var units = 1
      exportToSheet.getRange(MultiOutputRow,MultiunitsIndex).setValue(units);
      var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+"-"+ units;
      exportToSheet.getRange(MultiOutputRow, exportEndColIndex).setValue(exportcell);

      var marginperunit = exportToSheet.getRange(exportToSheet.getLastRow(),MarginPerUnitIndex).getValue();
      //2~10 ROWS OF MULTI QUOTE
      for (var i=1; i<=9; i++){
        //First 4 columns (DATES, CLIENT, LINK, PRODUCT CHINESE NAME)
        var exportColumnsNum = ChinesenameIndex-DateIndex+1;
        var copyRange = exportFromSheet.getRange(row, 1, 1, exportColumnsNum);  
        copyRange.copyTo(exportToSheet.getRange(exportToSheet.getLastRow()+1, DateIndex));

        //Second part -calculated price (COG, WEIGHT,MIN,GOAL,ACTUAL MARGIN)
        for (var k = ChinesenameIndex+1; k<=ChinesenameIndex+2; k++){
          var baseInfo = exportToSheet.getRange(exportToSheet.getLastRow()-i,k).getValue();
          exportToSheet.getRange(exportToSheet.getLastRow(),k).setValue(baseInfo+baseInfo*i);
        }
        for (var m = MinIndex; m<=MarginPerUnitIndex-1; m++){
          exportToSheet.getRange(exportToSheet.getLastRow(),m).setValue(exportToSheet.getRange(exportToSheet.getLastRow()-i,m).getValue()+marginperunit*i);
        }
        //Margin Per Unit remain same
        exportToSheet.getRange(exportToSheet.getLastRow(),MarginPerUnitIndex).setValue(marginperunit);

        //Third Part -Units
        exportToSheet.getRange(exportToSheet.getLastRow(),MarginPerUnitIndex+1).setValue(i+1);

        //Forth Part - static copy paste (NOTES, YUN BASIC, YUN VOL REG)
        for (var n = NotesIndex; n<= exportEndColIndex-1; n++){
          var Toheader = exportToSheet.getRange(1,n).getValue();                                
          if (Toheader in headerMap == true){
            var setcell = exportFromSheet.getRange(row,headerMap[Toheader]).getValue();     
            exportToSheet.getRange(exportToSheet.getLastRow(), ToheaderMap[Toheader]).setValue(setcell);
          }
        }
        //Last Part - update Sku name with new units (ENGLISH SKU NAME)
        var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+"-"+ (i+1);
        exportToSheet.getRange(exportToSheet.getLastRow(), exportEndColIndex).setValue(exportcell);
      } 

    }

  }
}




function QuoteByColumn(){ 
  var app = SpreadsheetApp;
  var exportFromSheet = app.getActiveSpreadsheet().getSheetByName("Quote Sheet");
  var lastrow = exportFromSheet.getLastRow();
  var lastcol = exportFromSheet.getLastColumn();
  var exportToSheet = app.getActiveSpreadsheet().getSheetByName("Quote Database");
  var totalcolumn = exportToSheet.getLastColumn();
  var totalrow = exportToSheet.getLastRow();

  for (var c = 1; c <= totalcolumn; c++){
    if (exportToSheet.getRange(1,c).getValue() == "English Sku name"){
      var exportEndColIndex = c;
    }
  }

  //store headers' index in "Quote sheet" in a hashtable for future lookup
  var ToheaderMap={};
  for (var m = 1; m<=totalcolumn;m++){
    var toheader = exportToSheet.getRange(1,m).getValue();
    ToheaderMap[toheader]=m;
  }
  //store headers' index in "Quote sheet" in a hashtable for future lookup
  var headerMap = {}; 
  for (var col = 1; col <= lastcol; col++){
    var header = exportFromSheet.getRange(1,col).getValue();
    headerMap[header] = col;
    }

  for (var row = 2; row <= lastrow; row++){
    //REGULAR
    //Loop columns in TOSheet to check if exist in 'headerMap' in FROMSheet then export cell by cell
    var RegularOutputRow = exportToSheet.getLastRow()+1;
    for (var k = 1; k <= totalcolumn; k++){
      if (headerMap[exportToSheet.getRange(1,k).getValue()]!=null){
        var cell = exportFromSheet.getRange(row,headerMap[exportToSheet.getRange(1,k).getValue()]).getValue();
        exportToSheet.getRange(RegularOutputRow,k).setValue(cell);
      }
    }

    //UPSELL
    if ((exportFromSheet.getRange(row,headerMap["Upsell quote?"]).getValue()) == "Yes"){
      var UpsellOutputRow = exportToSheet.getLastRow()+1;
      //First half
      for (var k = 1; k <= totalcolumn; k++){
        if (headerMap[exportToSheet.getRange(1,k).getValue()]!=null){
          var cell = exportFromSheet.getRange(row,headerMap[exportToSheet.getRange(1,k).getValue()]).getValue();
          exportToSheet.getRange(UpsellOutputRow,k).setValue(cell);
        }
      }
      //Set 'English SKU Name' column to name+upsell 
      var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+" "+"Upsell";
      exportToSheet.getRange(UpsellOutputRow,exportEndColIndex).setValue(exportcell);
    }

    //MULTISELL
    if ((exportFromSheet.getRange(row,headerMap["Multi quote?"]).getValue()) == "Yes"){
      var MultiOutputRow = exportToSheet.getLastRow()+1;
      //Loop columns in TOSheet to check if exist in 'headerMap' in FROMSheet and export cell by cell from 'Date'~('SKU NAME'-1)
      for (var k = 1; k <= totalcolumn; k++){
        if (headerMap[exportToSheet.getRange(1,k).getValue()]!=null){
          var cell = exportFromSheet.getRange(row,headerMap[exportToSheet.getRange(1,k).getValue()]).getValue();
          exportToSheet.getRange(MultiOutputRow,k).setValue(cell);
        }
      }
      //FIRST ROW OF MULTI QUOTE 'English SKU Name' column to name-1
      var units = 1
      exportToSheet.getRange(MultiOutputRow,ToheaderMap["Multi quote units"]).setValue(units);
      var exportcell = exportFromSheet.getRange(row,headerMap["English Sku name"]).getValue()+"-"+ units;
      exportToSheet.getRange(MultiOutputRow, ToheaderMap["English Sku name"]).setValue(exportcell);
  
      //SECOND ROW ~ 10th ROW
      var marginperunit = exportToSheet.getRange(exportToSheet.getLastRow(),ToheaderMap["Margin per unit"]).getValue();
      for (var i = 1; i<= 9; i++){
        var currentrow = exportToSheet.getLastRow()+1;
        for (var col = 1; col <= totalcolumn; col++){
          var colheader = exportToSheet.getRange(1,col).getValue();
          if (headerMap[colheader]!=null){
            if (colheader == "产品成本 COG"){
              var basecog = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(basecog+basecog*i);
            }
            else if (colheader == "重量 Weight"){
              var baseweight = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(baseweight+baseweight*i);         
            }
            else if (colheader =="最低利润 min."){
              var basemin = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(basemin+marginperunit*i);
            }
            else if (colheader == "理想利润 Goal"){
              var basegoal = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(basegoal+marginperunit*i);
            }
            else if (colheader == "Actual Margin"){
              var basemargin = exportToSheet.getRange(MultiOutputRow,ToheaderMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(basemargin+marginperunit*i);
            }
            else if (colheader == "Margin per unit"){
              exportToSheet.getRange(currentrow,ToheaderMap[colheader]).setValue(marginperunit);
            }
            else if (colheader == "English Sku name"){
              var exportcell = exportFromSheet.getRange(row,headerMap[colheader]).getValue()+"-"+ (i+1);
              exportToSheet.getRange(currentrow, ToheaderMap[colheader]).setValue(exportcell);          
            }
            else {
              var cell = exportFromSheet.getRange(row,headerMap[colheader]).getValue();
              exportToSheet.getRange(currentrow,col).setValue(cell);
            }
          }
        }
        //Column 'Multi quote units' adds 1
        exportToSheet.getRange(currentrow,ToheaderMap["Multi quote units"]).setValue(i+1);
      }
    }
  }
}


  