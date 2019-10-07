// Original script by Kirinya for this spreadsheet: https://docs.google.com/spreadsheets/d/1EOWc7kf9TdyvT31VfzlY284udUNOrtz0uyRtQ2t4MHY/
// Slightly edited by Mio Winter for this specific spreadsheet.
//
// In order to make this script public while at the same time keeping my API key private, I made a private spreadsheet
// that contains my API key in cell A1, and then I give this instance of the script permision to fetch that value.
// If you want to copy this script and use it for your own purposes, you need to replace the spreadsheet link below with
// your own, or alternatively simply set the 'OsuApiKey' variable to a string containing your osu! API key.
// Get your API key here: https://osu.ppy.sh/p/api
//  var osuApiSpreadsheet = "https://docs.google.com/spreadsheets/d/1qOH-ELkgIn_LJ45hd95wHXBXEgMaZoGc_gzppVInAgI/";
//  var osuApiKey = SpreadsheetApp.openByUrl(osuApiSpreadsheet).getActiveSheet().getRange(1,1).getValue();
var osuApiKey = '86f22be6ca202395194afd9d2e61eb973ff56d6d'; // Currently using Mio Winter's API key.

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


var LAST_COLUMN = OSU_COUNTRY_COLUMN;
var STARTING_COLUMN = 2;
var STARTING_ROW = 3;

var USER_HYPERLINK = '=HYPERLINK("https://osu.ppy.sh/u/PLACEHOLDER_USER_ID","PLACEHOLDER_USER_NAME")';

function updateAllPlayerData() {

  if (osuApiKey) {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Players');
    var numRows = (sheet.getLastRow() - STARTING_ROW + 1);
    var data = sheet.getRange(STARTING_ROW, PLAYER_COLUMN, numRows, LAST_COLUMN);
    var playerColumn = sheet.getRange(STARTING_ROW, PLAYER_COLUMN, numRows, 1);
    var playerLinks = playerColumn.getFormulas();
    var playerNames = playerColumn.getValues();
    var lastUpdatedColumn = sheet.getRange(STARTING_ROW, LAST_UPDATED_COLUMN, numRows, 1);
    
    for (var i = 1; i <= playerNames.length; i++) {
      var playerCell = data.getCell(i, PLAYER_COLUMN - STARTING_COLUMN + 1);
      var osuStdRankCell = data.getCell(i, OSU_STD_RANK_COLUMN - STARTING_COLUMN + 1);
      var osuStdPPCell = data.getCell(i, OSU_STD_PP_COLUMN - STARTING_COLUMN + 1);
      var osuTaikoRankCell = data.getCell(i, OSU_TAIKO_RANK_COLUMN - STARTING_COLUMN + 1);
      var osuTaikoPPCell = data.getCell(i, OSU_TAIKO_PP_COLUMN - STARTING_COLUMN + 1);
      var osuCatchRankCell = data.getCell(i, OSU_CATCH_RANK_COLUMN - STARTING_COLUMN + 1);
      var osuCatchPPCell = data.getCell(i, OSU_CATCH_PP_COLUMN - STARTING_COLUMN + 1);
      var osuManiaRankCell = data.getCell(i, OSU_MANIA_RANK_COLUMN - STARTING_COLUMN + 1);
      var osuManiaPPCell = data.getCell(i, OSU_MANIA_PP_COLUMN - STARTING_COLUMN + 1);
      var osuCountryCell = data.getCell(i, OSU_COUNTRY_COLUMN - STARTING_COLUMN + 1);
      var lastUpdatedCell = lastUpdatedColumn.getCell(i, 1);
      
      var formula = playerLinks[i-1][0];
      var enteredName = playerNames[i-1][0];

      if (enteredName) {
        var userID = getUserId(formula, enteredName);
        updatePlayerData(userID, playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, lastUpdatedCell);
      }
    }
    
  } else {
    throw("Error: API key not found");
  }
}

function updatePlayerData(userID, playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, lastUpdatedCell) {
  var today;
  
  if (!userID) { // invalid user ID (name-lookup failed)
    colorCells(playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, "red");
    return;
  }
  
  var jsonStd = getUserData(userID, 0);
  var jsonTaiko = getUserData(userID, 1);
  var jsonCatch = getUserData(userID, 2);
  var jsonMania = getUserData(userID, 3);
  
  if (jsonStd === undefined) { // player not found
    colorCells(playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, "red");
    playerCell.setComment("Player not found.");
    return;
  } else {
    var playerName = jsonStd.username;
    var osuCountry = jsonStd.country;
    var osuStdRank = jsonStd.pp_rank;
    var osuStdPP = jsonStd.pp_raw;
    var osuTaikoRank = jsonTaiko.pp_rank;
    var osuTaikoPP = jsonTaiko.pp_raw;
    var osuCatchRank = jsonCatch.pp_rank;
    var osuCatchPP = jsonCatch.pp_raw;
    var osuManiaRank = jsonMania.pp_rank;
    var osuManiaPP = jsonMania.pp_raw;
    
    // Colour cells grey if player is inactive
    if (!(osuStdPP > 0) && !(osuTaikoPP > 0) && !(osuCatchPP > 0) && !(osuManiaPP > 0)) { // player is inactive in all modes
      colorCells(playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, "lightgrey");
      playerCell.setComment("Player is inactive.\n");

    } else {
      // Remove colouring in case player was previously inactive and has returned
      colorCells(playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, null);
    }
    
    // Update player data
    updateRanking(osuStdRankCell, osuStdRank, osuStdPPCell, osuStdPP, osuTaikoRankCell, osuTaikoRank, osuTaikoPPCell, osuTaikoPP, osuCatchRankCell, osuCatchRank, osuCatchPPCell, osuCatchPP, osuManiaRankCell, osuManiaRank, osuManiaPPCell, osuManiaPP, osuCountryCell, osuCountry);
    updatePlayerName(userID, playerCell, playerName);
    
    // Write in date of update as today.
    today = new Date(); // This gets a timestamp of when the variable is defined, which is sometime during the execution of script.
    lastUpdatedCell.setValue(today);
  }
}

function getUserData(userID, mode) { // use mode = 0 for osu!std, 1 for osu!taiko, 2 for osu!catch and 3 for osu!mania
  var cache = CacheService.getScriptCache();
  var cached = cache.get(userID + " " + mode);
  if (cached) {
    Logger.log("Cached content: " + cached);
    return JSON.parse(cached)[0];
  }
  
  var response = UrlFetchApp.fetch("https://osu.ppy.sh/api/get_user\?k\=" + osuApiKey + "\&u\=" + userID + "\&type\=id\&m\=" + mode + "\&limit\=1");
    // looks like & gets replaced with &amp; .... maybe cause of [] empty response
  while (!response.getContentText()[0]) {
    Logger.log("Failed response. Trying again for " + userID);
    response = UrlFetchApp.fetch("https://osu.ppy.sh/api/get_user\?k\=" + osuApiKey + "\&u\=" + userID + "\&type\=id\&m\=" + mode + "\&limit\=1");
  }
    //sometimes response is empty. probably too many API calls in a short time.
  var content = response.getContentText();
  if (content) {
    Logger.log("New REST content: " + content);
    cache.put(userID + " " + mode, content, 21600); // cache result for 6 hours.
    return JSON.parse(content)[0];
  } else {
  return null;
  }
}

function getUserId(hyperlink, enteredName) {
  if (!hyperlink) { //profile link doesn't exist
    var response = UrlFetchApp.fetch("https://osu.ppy.sh/api/get_user\?k\=" + osuApiKey + "\&u\=" + enteredName + "\&type\=string\&m\=0\&limit\=1");
    var json = JSON.parse(response.getContentText())[0];
    if (json && json.user_id) {
      return json.user_id;
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast("Player '" + enteredName + "' not found.", "Input error!");
      return null;
    }
  } else {
    return getUserIdFromHyperlinkText(hyperlink);
  }
}

function getUserIdFromHyperlinkText(hyperlink) {
  var regExpForUserID = new RegExp("[0-9]+");
  var userID = regExpForUserID.exec(hyperlink)[0];
  if (userID) { //hyperlink contained ID number
    return userID;
  } else { //hyperlink was not valid ID number
    return null;
  }
}

function updateRanking(osuStdRankCell, osuStdRank, osuStdPPCell, osuStdPP, osuTaikoRankCell, osuTaikoRank, osuTaikoPPCell, osuTaikoPP, osuCatchRankCell, osuCatchRank, osuCatchPPCell, osuCatchPP, osuManiaRankCell, osuManiaRank, osuManiaPPCell, osuManiaPP, osuCountryCell, osuCountry) {
  osuStdRankCell.setValue(osuStdRank);
  osuStdPPCell.setValue(osuStdPP);
  osuTaikoRankCell.setValue(osuTaikoRank);
  osuTaikoPPCell.setValue(osuTaikoPP);
  osuCatchRankCell.setValue(osuCatchRank);
  osuCatchPPCell.setValue(osuCatchPP);
  osuManiaRankCell.setValue(osuManiaRank);
  osuManiaPPCell.setValue(osuManiaPP);
  osuCountryCell.setValue(osuCountry);
  
  if (!(osuStdPP > 300)) { // If user does not have over 300 PP in a game mode, I assume they're inactive in that game mode.
    osuStdRankCell.setBackground("lightgrey");
    osuStdPPCell.setBackground("lightgrey");
  }
  if (!(osuTaikoPP > 300)) {
    osuTaikoRankCell.setBackground("lightgrey");
    osuTaikoPPCell.setBackground("lightgrey");
  }
  if (!(osuCatchPP > 300)) {
    osuCatchRankCell.setBackground("lightgrey");
    osuCatchPPCell.setBackground("lightgrey");
  }
  if (!(osuManiaPP > 300)) {
    osuManiaRankCell.setBackground("lightgrey");
    osuManiaPPCell.setBackground("lightgrey");
  }
}

function updatePlayerName(userID, playerCell, playerName) {
  var oldPlayerName;
  
  //This will write the old name into a 'note' in the cell.
  oldPlayerName = playerCell.getValue();
  if (!(oldPlayerName === playerName)) {
    var comment = playerCell.getComment();
    playerCell.setComment(comment + oldPlayerName + "\n");
  }
  
  // Update player name
  playerCell.setFormula(USER_HYPERLINK.replace("PLACEHOLDER_USER_ID", userID).replace("PLACEHOLDER_USER_NAME", playerName)).setFontWeight('bold').setFontLine('underline');
}

function colorCells(playerCell, osuStdRankCell, osuStdPPCell, osuTaikoRankCell, osuTaikoPPCell, osuCatchRankCell, osuCatchPPCell, osuManiaRankCell, osuManiaPPCell, osuCountryCell, color) {
  playerCell.setBackground(color);
  osuStdRankCell.setBackground(color);
  osuStdPPCell.setBackground(color);
  osuTaikoRankCell.setBackground(color);
  osuTaikoPPCell.setBackground(color);
  osuCatchRankCell.setBackground(color);
  osuCatchPPCell.setBackground(color);
  osuManiaRankCell.setBackground(color);
  osuManiaPPCell.setBackground(color);
  osuCountryCell.setBackground(color);
}
