//Sheet link: Sheet link: https://docs.google.com/spreadsheets/d/1trJ-2F-dkZaZes8EUBns2ift4JzB2NrEwIatAt6_0fI/edit#gid=268970711

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
