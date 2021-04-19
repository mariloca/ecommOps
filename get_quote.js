
//Sheet link: https://docs.google.com/spreadsheets/d/1Z1WU8w0D7gWhT4JhpNs9B3dNFI8xI61i4S8lRTMhZ0k/edit#gid=119345874

/**
 * Get shipping quote
 *
 * @param SK
 * @param Number of Units
 * @param Country
 * @param Shipping Method
 * @return Shipping cost
 * @customfunction
 */
function getShippingQuote(sku,number,country,line){
  
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getActiveSheet();
  var priceSheet = app.getActiveSpreadsheet().getSheetByName("Product Prices");

  /* Input example:
  sku = "Bands";
  number = "1";
  country = "Ireland";
  line = "yunexpress";
  */
  
  //concatenate country with lines and sku with unit number
  country = properCase(country);
  var shipping = country + " " + line;
  var target = sku + "-" + number;
  var targetRow = priceSheet.getLastRow();
  var targetColumn = priceSheet.getLastColumn();
  Logger.log(shipping)
  Logger.log(target)
  
  var rowIndex = 0 
  var colIndex = 0
  
  //find row index and column index of the input sku for input country shipping lines
  for (var r = 1; r<=targetRow; r++){
    if (target === priceSheet.getRange(r, 1).getValue()){
      rowIndex = r;
    }  
  }
  for (var c = 1; c <= targetColumn; c++){
    if (shipping === priceSheet.getRange(1, c).getValue()){
      colIndex = c;
    }
  }
  
  //if not find column index for shipping line, then use 'other' for shipping line
  if (colIndex == 0){
      colIndex = targetColumn;
    }
  
  var output = priceSheet.getRange(rowIndex, colIndex).getValue();
  Logger.log(shipping);
  return output
}


/**
 * Capitalized first letter of a word
 *
 * @param {word} input A word to capitalize.
 * @return Capitalized word.
 * @customfunction
 */
function properCase(phrase) {
  // /\b(\w)/g selects the first letter of each word. /\B(\w)/g selects all the characters in each word except the first letter.
  var regFirstLetter = /\b(\w)/g;
  var regOtherLetters = /\B(\w)/g;
  function capitalize(firstLetters) {
    return firstLetters.toUpperCase();
  }
  function lowercase(otherLetters) {
    return otherLetters.toLowerCase();
  }
  var capitalized = phrase.replace(regFirstLetter, capitalize);
  var proper = capitalized.replace(regOtherLetters, lowercase);

  return proper;
}