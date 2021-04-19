//Sheet link: https://docs.google.com/spreadsheets/d/1trJ-2F-dkZaZes8EUBns2ift4JzB2NrEwIatAt6_0fI/edit#gid=268970711


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
