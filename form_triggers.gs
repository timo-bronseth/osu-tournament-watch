// Made by Mio Winter.
// 23/05/2018.

// This function is triggered by a form submission, and is then fed an object 'e' containing the form submission details
// Documentation about trigger event objects here: https://developers.google.com/apps-script/guides/triggers/events
function formTrigger (e) {
  var sheet;
  
  // The called function deletes row if applicable, and returns 1 if the row was deleted, 0 if not.
  // Works for both forms.
  if (deleteRowIfMarked(e)) { 
    return; // Stop the function if the row was deleted.
  }
  
  // I have two forms linked to this spreadsheet, and this function runs whenever either of the forms are submitted.
  // I only want to add the EditResponseUrl when the tournaments form is submitted or edited, so I only run the following
  // code if the response contains a field named 'Tournament name'.
  if (e.namedValues['Tournament name']) {
    addEditUrl(e);
    addTournamentFormulas(e);
    SpreadsheetApp.flush(); // Makes sure the previous steps are completed before continuing on to next step.
    reorganiseTournamentsSheet();
  } else { // And the following code runs when the 'newsletter' form is submitted.
    addPlayerFormulas(e);
    callUpdatePlayerDataFunction(e);
  }
}


function callUpdatePlayerDataFunction (e) {
  var sheet, rowIndex, activeRow, playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, lastUpdatedCell, userID;
  
  var PLAYER_COLUMN = 2;
  var OSU_STD_RANK_COLUMN = 3;
  var OSU_STD_PP_COLUMN = 4;
  var OSU_TAIKO_RANK_COLUMN = 5;
  var OSU_TAIKO_PP_COLUMN = 6;
  var OSU_CATCH_RANK_COLUMN = 7;
  var OSU_CATCH_PP_COLUMN = 8;
  var OSU_MANIA_RANK_COLUMN = 9;
  var OSU_MANIA_PP_COLUMN = 10;
  var OSU_COUNTRY_COLUMN = 11;
  var LAST_UPDATED_COLUMN = 25;
  var LAST_COLUMN = LAST_UPDATED_COLUMN;
  
  sheet = e.range.getSheet();
  rowIndex = e.range.getRowIndex(); // Gets the row index of the cells that the submissions were entered into.
  activeRow = sheet.getRange(rowIndex, 1, 1, LAST_COLUMN);
  playerCell = activeRow.getCell(1, PLAYER_COLUMN);
  osuStdRankCell = activeRow.getCell(1, OSU_STD_RANK_COLUMN);
  osuStdPPCell = activeRow.getCell(1, OSU_STD_PP_COLUMN);
  osuTaikoRankCell = activeRow.getCell(1, OSU_TAIKO_RANK_COLUMN);
  osuTaikoPPCell = activeRow.getCell(1, OSU_TAIKO_PP_COLUMN);
  osuCatchRankCell = activeRow.getCell(1, OSU_CATCH_RANK_COLUMN);
  osuCatchPPCell = activeRow.getCell(1, OSU_CATCH_PP_COLUMN);
  osuManiaRankCell = activeRow.getCell(1, OSU_MANIA_RANK_COLUMN);
  osuManiaPPCell = activeRow.getCell(1, OSU_MANIA_PP_COLUMN);
  osuCountryCell = activeRow.getCell(1, OSU_COUNTRY_COLUMN);
  lastUpdatedCell = activeRow.getCell(1, LAST_UPDATED_COLUMN);
  
  userID = getUserId(playerCell.getFormula(), playerCell.getValue());

  updatePlayerData(userID, playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, lastUpdatedCell);
}

function deleteRowIfMarked(e) {
  var rowIndex, sheet, matchingSheet;
  
  if (e.namedValues['Do you wish to delete this entry?'] == 'Yes') {
    sheet = e.range.getSheet();
    rowIndex = e.range.getRowIndex();
    
    sheet.deleteRow(rowIndex);
    
    if (e.namedValues['Tournament name']) {
      matchingSheet = SpreadsheetApp.getActive().getSheetByName('TournamentPlayerMatching');
      matchingSheet.deleteRow(rowIndex);
      // If you delete a row in the 'Tournaments' sheet, the row in the matchingSheet that referred to the deleted row, will get "#REF" errors.
      // So you need to delete the rows in both sheets if you delete one in the 'Tournaments' sheet.
    }
    
    Logger.log("Row deleted: " + rowIndex);
    return 1; // Tells the caller that the row was deleted.
  } else {
    return 0;
  } 
}

function addEditUrl(e) {
  var form, sheet, rowIndex, timestamp, formResponses, urlCell;
  
  var EDITURL_COLUMN = 21;
  
  form = FormApp.openById('19nwUAJGa8A902cjkpmmM_ECgIb6BwDsRAKwfwz4bQxc'); // This is the tournaments form.
  sheet = e.range.getSheet();
  rowIndex = e.range.getRowIndex(); // Gets the row index of the cells that the submissions were entered into.
  timestamp = sheet.getRange(rowIndex, 1).getValue();
  formResponses = form.getResponses(timestamp); // Gets an array of all the responses with that timestamp.
  urlCell = sheet.getRange(rowIndex, EDITURL_COLUMN); // The cell into which the url will be written
  
  // Set cell value and formatting
  urlCell.setFormula('=HYPERLINK("' + formResponses[0].getEditResponseUrl() + '", "EDIT")');
  urlCell.setBackground('black').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
}

function addTournamentFormulas(e) {
  var sheet, rowIndex, hyperlinkCell, divisionsCell, maxRankCell, minRankCell;
  
  var HYPERLINK_COLUMN = 2;
  var DIVISIONS_COLUMN = 12;
  var MAXRANK_COLUMN = 13;
  var MINRANK_COLUMN = 14;
  
  sheet = e.range.getSheet();
  rowIndex = e.range.getRowIndex();
  hyperlinkCell = sheet.getRange(rowIndex, HYPERLINK_COLUMN);
  divisionsCell = sheet.getRange(rowIndex, DIVISIONS_COLUMN);
  maxRankCell = sheet.getRange(rowIndex, MAXRANK_COLUMN);
  minRankCell = sheet.getRange(rowIndex, MINRANK_COLUMN);
  
  hyperlinkCell.setFormula(
    '=IFERROR(HYPERLINK(D' + rowIndex + ', C' + rowIndex + '))'
  );
  
  hyperlinkCell.setFontWeight('bold');
  
  divisionsCell.setFormula(
    '=IF(ISBLANK(J' + rowIndex + '), IFERROR(1/0), COUNTA(SPLIT(J' + rowIndex + ',",")))'
  );
  
  // Note that I had to escape the "\" character in the following setFormulas, so they look like "\\+" and will appear "\+" in the formula in the spreadsheet.
  maxRankCell.setFormula(
    '=IFERROR(IF(ISBLANK(C' + rowIndex + '), IFERROR(1/0), IF(ISBLANK(J' + rowIndex + '), 0, MIN(Split(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(J' + rowIndex + ', "(?i)k", "000")," ", ""), "-", ","), "\\+", ""), ",", false)))))'
  );
  
  minRankCell.setFormula(
    '=IFERROR(IF(ISBLANK(C' + rowIndex + '), IFERROR(1/0), IF(ISBLANK(J' + rowIndex + '), 99999999, IF(OR(REGEXMATCH(J' + rowIndex + ',"\\+"), REGEXMATCH(J' + rowIndex + ',"∞")),99999999,MAX(Split(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(J' + rowIndex + ', "(?i)k", "000")," ", ""), "-", ","), "\\+", ""), ",", false))))))'
  );
}

function addPlayerFormulas(e) {
  var sheet, rowIndex, matchesCell, otherTourniesCell;
  
  var STARTING_ROW = 3;
  var MATCHES_COLUMN = 20;
  var OTHERTOURNIES_COLUMN = 21;
  
  sheet = e.range.getSheet();
  rowIndex = e.range.getRowIndex();
  matchesCell = sheet.getRange(rowIndex, MATCHES_COLUMN);
  otherTourniesCell = sheet.getRange(rowIndex, OTHERTOURNIES_COLUMN);
  
  matchesCell.setFormula(
    '=IFERROR(IF(NOT(ISBLANK(L' + rowIndex + ')), JOIN(char(10), FILTER(TournamentPlayerMatching!A$3:A, REGEXMATCH(TournamentPlayerMatching!H$3:H, L' + rowIndex + '), NOT(ISBLANK(TournamentPlayerMatching!H$3:H)))), IFERROR(1/0)))'
  );
  
  otherTourniesCell.setFormula(
    '=IFERROR(IF(ISBLANK(B' + rowIndex + '), IFERROR(1/0), TEXTJOIN(char(10), TRUE, FILTER(TRANSPOSE(SPLIT(TournamentPlayerMatching!$K$3, char(10))), ISNA(MATCH(TRANSPOSE(SPLIT(TournamentPlayerMatching!$K$3, char(10))), TRANSPOSE(SPLIT(T' + rowIndex + ', char(10))),0))))))'
  );
}
