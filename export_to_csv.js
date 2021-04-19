//Sheet link: https://docs.google.com/spreadsheets/d/1Z1WU8w0D7gWhT4JhpNs9B3dNFI8xI61i4S8lRTMhZ0k/edit#gid=119345874

function saveAsCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName('Client quotes');
  // create a folder from the name of the spreadsheet
  var formattedDate = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
  var folder = DriveApp.createFolder(ss.getName()+ '_csv_' + formattedDate);
  //var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + formattedDate);
  // append ".csv" extension to the sheet name
  fileName = sheet.getName() + " " + formattedDate + ".csv" ;
  // convert all available sheet data to csv format
  var csvFile = convertRangeToCsvFile_(fileName, sheet);
  // create a file in the Docs List with the given name and the csv data
  var file = folder.createFile(fileName, csvFile);
  //File downlaod
  var downloadURL = file.getDownloadUrl().slice(0, -8);
  showurl(downloadURL);
  
}

function showurl(downloadURL) {
  //Change what the download button says here
  var link = HtmlService.createHtmlOutput('<a href="' + downloadURL + '">Click here to download</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Your CSV file is ready!');
}

function convertRangeToCsvFile_(csvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
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
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}