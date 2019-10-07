// Made by Mio Winter.
// 11/05/2018.

function reorganiseTournamentsSheet() {
  var sheet, numRows;
  
  var STARTING_ROW = 3;
  var NAMES_COLUMN = 3;
  var DATES_COLUMN = 8;
  var DELETE_COLUMN = 18;
  var ALWAYSOPEN_COLUMN = 20;
  var LAST_COLUMN = 21;
  
  sheet = SpreadsheetApp.getActive().getSheetByName('Tournaments');
  numRows = (getFirstEmptyRow(sheet.getRange(STARTING_ROW, NAMES_COLUMN, sheet.getLastRow())) - STARTING_ROW); // Getting the value here, so that I don't have several calls to this function for each of the functions below.

  colourIfStillOpen(sheet, STARTING_ROW, numRows, DATES_COLUMN, ALWAYSOPEN_COLUMN, LAST_COLUMN);
  SpreadsheetApp.flush(); // Makes sure the previous step is completed before continuing on to next step.
  sortByDate(sheet, STARTING_ROW, DATES_COLUMN, numRows, LAST_COLUMN);
}

function colourIfStillOpen(sheet, STARTING_ROW, numRows, DATES_COLUMN, ALWAYSOPEN_COLUMN, LAST_COLUMN) {
  var dates, alwaysOpenValues, range, dateCell, today, compareDate, dateInFuture, row;
  
  dates = sheet.getRange(STARTING_ROW, DATES_COLUMN, numRows, 1).getValues();
  alwaysOpenValues = sheet.getRange(STARTING_ROW, ALWAYSOPEN_COLUMN, numRows, 1).getValues();
  
  today = new Date();
    
  for (row = 0; row < numRows; row++) {
    compareDate = dates[row];
    range = sheet.getRange(STARTING_ROW+row, 1, 1, LAST_COLUMN-1);
    dateCell = sheet.getRange(STARTING_ROW+row, DATES_COLUMN, 1);
    
    if (alwaysOpenValues[row] == 'Yes') {
      Logger.log("Colouring cells green at row: " + (STARTING_ROW+row));
      dateCell.setValue(" (always open)");
      range.setBackgroundRGB(0,255,0);
      
    } else if ((compareDate.toString().length == 3) || (compareDate.toString().length == 12)) {
      Logger.log("Row already set at row: " + (STARTING_ROW+row) + " Length: " + compareDate.toString().length);
      
    } else {
      compareDate = new Date(compareDate); // Need to cast it in date format.
      dateInFuture = isDateInFuture(today.valueOf(), compareDate.valueOf());
      
      if (dateInFuture == false) {
        Logger.log("Colouring cells grey at row: " + STARTING_ROW+row);
        dateCell.setValue("(" + Utilities.formatDate(compareDate, 'GMT', 'dd/MM/yyyy') + ")");
        range.setBackground('lightgrey');
      } else if (dateInFuture == true) {
        Logger.log("Colouring cells green at row: " + (STARTING_ROW+row));
        range.setBackgroundRGB(0,255,0);
      } else {
        Logger.log("Empty date at row: " + (STARTING_ROW+row));
        dateCell.setValue("(?)");
        range.setBackground(null);
      }
    }
  }
}
function sortByDate(sheet, STARTING_ROW, DATES_COLUMN, numRows, LAST_COLUMN) {
  var range;
  
  range = sheet.getRange(STARTING_ROW, 1, numRows, LAST_COLUMN);
  range.sort(DATES_COLUMN);
}

function getFirstEmptyRow (range) { // Takes a range with one column only
  var values, i;
  
  values = range.getValues();
  i = 0;
  while (values[i].toString()) i++;
  return range.getCell(i+1, 1).getRowIndex();
}

function isDateInFuture(todayValue, compareDateValue) { // Use date.valueOf() as arguments
  var MS_PER_HOUR;
  
  // Date().valueOf() in JavaScript/GAS returns the number of milliseconds since midnight January 1st, 1970.
  // Date(<some date>) returns the number of milliseconds between midnight January 1st, 1970,
  // and the START of <some date>. So if I want to number of ms to the END of <some date>, I need to
  // take <some date> and add the number of milliseconds in 24 hours.
  // By the way, my spreadsheet uses GMT time, but if you copy the script it will use your local settings.
  MS_PER_HOUR = 60 * 60 * 1000;
  
  if (todayValue < compareDateValue + MS_PER_HOUR*24) {
    return true;// return true if date is in future.
  } else if (todayValue > compareDateValue + MS_PER_HOUR*24) {
    return false;//return false if date is not in future.
  } else {
    return null;// return null if variable is not a date.
  }
}

function has2WeeksPassed(compareDate) {
  var today, MS_PER_HOUR;
  
  if (compareDate.toString().length == 0) {return true}; // Return true if compareDate is empty, which means the script has never sent the player an email before.
  
  // Date().valueOf() in JavaScript/GAS returns the number of milliseconds since midnight January 1st, 1970.
  // Date(<some date>) returns the number of milliseconds between midnight January 1st, 1970,
  // and the START of <some date>. So if I want to number of ms to the END of <some date>, I need to
  // take <some date> and add the number of milliseconds in 24 hours.
  // By the way, my spreadsheet uses GMT time, but if you copy the script it will use your local settings.
  MS_PER_HOUR = 60 * 60 * 1000;

  today = new Date().valueOf();
  compareDate = new Date(compareDate).valueOf(); // This also converts compareDate to date format, just in case input was a string.
  if (today > compareDate + MS_PER_HOUR*24*14) {
    return true;// return true if date is more than a week ago.
  } else if (today < compareDate + MS_PER_HOUR*24*7) {
    return false;//return false if date is less than a week ago.
  }
}

function clearComments () {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Tournaments');
  var lastRow = sheet.getLastRow();
  var COLUMN = 8;
  
  for (var row=1;row<lastRow;row++) {
    sheet.getRange(row, COLUMN).clearNote();
  }
}
