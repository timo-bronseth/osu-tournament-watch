// Made by Mio Winter.
// 21/05/2018.

function preEmailUpdates() {
  // I just add this script to this file because I'm lazy and don't want to change script files to run each function before I run sendEmail().
  updateAllPlayerData();
  reorganiseTournamentsSheet();
}


// The sendEmail() function should be triggered every day (1pm to 2pm), but I've deleted the trigger for now because my system cannot handle >200 subscribers.
// Takes more of Google's processing power than my free account lets me have.
function sendEmails() {
  var playerObjectList, emailCondition, dailyQuota, subject, htmlBody, i;
  
  storeTournyDataInCache();
  // Store tourny data in cache before sending emails, so that the script doesn't have to look those up in the spreadsheet for every email it sends.
  
  playerObjectList = getPlayerObjectList();
  // This returns a list of player objects that are filtered for players who last received email more than two weeks ago, and sorted according to who received mails last.
  // The 'sendEmails()' function is triggered at 00:00 every day, and will just send 0 emails the filtered-and-sorted playerObjectList has 0 players in it.
  
  for (i = 0; i < playerObjectList.length; i++) {
    // Update all player info and flush the spreadsheet, BEFORE sending emails, to make sure emails are sent with updated info.
    var userID = getUserId(playerObjectList[i].formula, playerObjectList[i].name);
    updatePlayerData(userID,
                     playerObjectList[i].playerNameCell,
                     playerObjectList[i].osuStdRankCell,
                     playerObjectList[i].osuStdPPCell,
                     playerObjectList[i].osuTaikoRankCell,
                     playerObjectList[i].osuTaikoPPCell,
                     playerObjectList[i].osuCatchRankCell,
                     playerObjectList[i].osuCatchPPCell,
                     playerObjectList[i].osuManiaRankCell,
                     playerObjectList[i].osuManiaPPCell,
                     playerObjectList[i].osuCountryCell,
                     playerObjectList[i].lastUpdatedCell);
  }
  
  SpreadsheetApp.flush(); // This makes sure all previous updates are applied before continuing to execute the next part of the script.
  
  for (i = 0; i < playerObjectList.length; i++) {

    // Only send email if either 1) they have not gotten an email before,
    // OR 2) the player's tournament matches are different from the last tournament matches they got a mail about,
    // OR 3) if they selected "No" to the "restrict updates?" question and their non-matching tournaments are different from the last non-matching tournaments they got mail about.
    // If the both their tournament matches and non-matches are identical, both of these conditions are false and they will not get an email.    
    emailCondition = (playerObjectList[i].lastMailDateCell.getValue().toString().length == 0) || 
                     (playerObjectList[i].lastMailMatches != playerObjectList[i].tournyMatches.toString()) ||
                     ((playerObjectList[i].restrictUpdates.toString() == 'No, send me emails about all tournament updates') && (playerObjectList[i].lastMailOthers != playerObjectList[i].otherTournies.toString()));
    Logger.log("EVALUATING PLAYER: " + playerObjectList[i].name + " Send email? " + emailCondition);

    if (emailCondition) {
      subject = getSubject(playerObjectList[i].name);
      htmlBody = getHtmlBody(playerObjectList[i].mail, playerObjectList[i].name, playerObjectList[i].tournyMatches.toString().split("\n"), playerObjectList[i].otherTournies.toString().split("\n"));
      MailApp.sendEmail(playerObjectList[i].mail, subject, 'placeholderBody', {htmlBody: htmlBody});
      playerObjectList[i].lastMailDateCell.setValue(new Date());
      playerObjectList[i].lastMailMatchesCell.setValue(playerObjectList[i].tournyMatches);
      playerObjectList[i].lastMailOthersCell.setValue(playerObjectList[i].otherTournies);
      Logger.log("Email sent to: " + playerObjectList[i].mail);
      dailyQuota = MailApp.getRemainingDailyQuota();
      Logger.log("Remaining daily quota: " + dailyQuota + '.');
      
      if (!(dailyQuota > 0)) {
        // Standard google accounts (like mine) are limited to sending 100 emails via scripts per day.
        // If the script still needs to send emails after RemainingDailyQuota == 0, it will break out of the loop.
        Logger.log("Script has exceeded its daily email quota limit. Breaking out of loop.");
        break;
      }
    }
  }
}

function getPlayerObjectList() {
  var sheet, lastRow, playerMailsColumn, playerMails, playerNames, playerFormulas, playerRestrictUpdates, playerTournyMatches, playerOtherTournies, playerLastMailMatchesColumn, playerLastMailMatches, playerLastMailOthersColumn, playerLastMailOthers, playerLastMailDateColumn, playerLastMailDate, playerNameColumn, osuStdRankColumn, osuStdPPColumn, osuTaikoRankColumn, osuTaikoPPColumn, osuCatchRankColumn, osuCatchPPColumn, osuManiaRankColumn, osuManiaPPColumn, osuCountryColumn, lastUpdatedColumn, playerObjectList, i;
  
  var MAIL_COLUMN = 12;
  var PLAYERNAME_COLUMN = 2;
  var RESTRICTUPDATES_COLUMN = 17;
  var MATCHES_COLUMN = 20;
  var OTHERTOURNIES_COLUMN = 21;
  var LASTMAILSENT_COLUMN = 22;
  var LASTMAILMATCHES_COLUMN = 23;
  var LASTMAILOTHERS_COLUMN = 24;
  var STARTING_ROW = 3;
  
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
  
  sheet = SpreadsheetApp.getActive().getSheetByName('Players');
  lastRow = sheet.getLastRow();
  playerMailsColumn = sheet.getRange(STARTING_ROW, MAIL_COLUMN, lastRow);
  playerMails = playerMailsColumn.getValues();
  playerNames = sheet.getRange(STARTING_ROW, PLAYERNAME_COLUMN, lastRow).getValues();
  playerFormulas = sheet.getRange(STARTING_ROW, PLAYERNAME_COLUMN, lastRow).getFormulas();
  playerRestrictUpdates = sheet.getRange(STARTING_ROW, RESTRICTUPDATES_COLUMN, lastRow).getValues(); // This references players' preferences about whether to update them on all tournaments or just those that match them.
  playerTournyMatches = sheet.getRange(STARTING_ROW, MATCHES_COLUMN, lastRow).getValues(); // This references the tournaments that matches each player's preferences and eligibility.
  playerOtherTournies = sheet.getRange(STARTING_ROW, OTHERTOURNIES_COLUMN, lastRow).getValues(); // This references all open tournaments EXCEPT the ones that match the player.
  playerLastMailMatchesColumn = sheet.getRange(STARTING_ROW, LASTMAILMATCHES_COLUMN, lastRow);
  playerLastMailMatches = sheet.getRange(STARTING_ROW, LASTMAILMATCHES_COLUMN, lastRow).getValues();
  playerLastMailOthersColumn = sheet.getRange(STARTING_ROW, LASTMAILOTHERS_COLUMN, lastRow);
  playerLastMailOthers = sheet.getRange(STARTING_ROW, LASTMAILOTHERS_COLUMN, lastRow).getValues();
  playerLastMailDateColumn = sheet.getRange(STARTING_ROW, LASTMAILSENT_COLUMN, lastRow);
  playerLastMailDate = sheet.getRange(STARTING_ROW, LASTMAILSENT_COLUMN, lastRow).getValues(); // This references the dates on which the last emails were sent to each player.
  
  playerNameColumn = sheet.getRange(STARTING_ROW, PLAYERNAME_COLUMN, lastRow);
  osuStdRankColumn = sheet.getRange(STARTING_ROW, OSU_STD_RANK_COLUMN, lastRow);
  osuStdPPColumn = sheet.getRange(STARTING_ROW, OSU_STD_PP_COLUMN, lastRow);
  osuTaikoRankColumn = sheet.getRange(STARTING_ROW, OSU_TAIKO_RANK_COLUMN, lastRow);
  osuTaikoPPColumn = sheet.getRange(STARTING_ROW, OSU_TAIKO_PP_COLUMN, lastRow);
  osuCatchRankColumn = sheet.getRange(STARTING_ROW, OSU_CATCH_RANK_COLUMN, lastRow);
  osuCatchPPColumn = sheet.getRange(STARTING_ROW, OSU_CATCH_PP_COLUMN, lastRow);
  osuManiaRankColumn = sheet.getRange(STARTING_ROW, OSU_MANIA_RANK_COLUMN, lastRow);
  osuManiaPPColumn = sheet.getRange(STARTING_ROW, OSU_MANIA_PP_COLUMN, lastRow);
  osuCountryColumn = sheet.getRange(STARTING_ROW, OSU_COUNTRY_COLUMN, lastRow);
  lastUpdatedColumn = sheet.getRange(STARTING_ROW, LAST_UPDATED_COLUMN, lastRow);
  
  playerObjectList = []; // Create an array of objects that contain information about each player, so that the array can be sorted according to who has waited longest since their last email.
  
  for (i = 0; i < (getFirstEmptyRow(playerMailsColumn) - STARTING_ROW); i++) {
    playerObjectList.push({
      mail: playerMails[i],
      name: playerNames[i],
      formula: playerFormulas[i],
      restrictUpdates: playerRestrictUpdates[i],
      tournyMatches: playerTournyMatches[i],
      otherTournies: playerOtherTournies[i],
      lastMailMatchesCell: playerLastMailMatchesColumn.getCell(i+1,1), // need to reference the cell itself so that it can be updated once a mail is sent to player.
      lastMailMatches: playerLastMailMatches[i],
      lastMailOthersCell: playerLastMailOthersColumn.getCell(i+1,1),
      lastMailOthers: playerLastMailOthers[i],
      lastMailDateCell: playerLastMailDateColumn.getCell(i+1, 1),
      lastMailDate: playerLastMailDate[i],
      
      playerNameCell: playerNameColumn.getCell(i+1, 1),
      osuStdRankCell: osuStdRankColumn.getCell(i+1, 1),
      osuStdPPCell: osuStdPPColumn.getCell(i+1, 1),
      osuTaikoRankCell: osuTaikoRankColumn.getCell(i+1, 1),
      osuTaikoPPCell: osuTaikoPPColumn.getCell(i+1, 1),
      osuCatchRankCell: osuCatchRankColumn.getCell(i+1, 1),
      osuCatchPPCell: osuCatchPPColumn.getCell(i+1, 1),
      osuManiaRankCell: osuManiaRankColumn.getCell(i+1, 1),
      osuManiaPPCell: osuManiaPPColumn.getCell(i+1, 1),
      osuCountryCell: osuCountryColumn.getCell(i+1, 1),
      lastUpdatedCell: lastUpdatedColumn.getCell(i+1, 1)
    })
  }
  
  playerObjectList = playerObjectList.filter(function (player) {
    return player.mail.toString().length > 0 && // filter out players who have nothing in the mail column.
           has2WeeksPassed(player.lastMailDate) && // filter out players who have received email within the last two weeks.
           player.formula.toString().length > 0; // filter out players who do not have updated player information (pp values and so on).
  });
  
  playerObjectList.sort(function(a, b) { // This sorts the list in descending order according to when emails were sent to each player.
    var dateA = new Date(a.lastMailDate).valueOf();
    var dateB = new Date(b.lastMailDate).valueOf();
    if (!dateA) {dateA = 0;} // dateA will be "NaN" for empty values, but I want it to be 0 so that it is sorted correctly.
    if (!dateB) {dateB = 0;}
    
    return ((dateA < dateB) ? -1 : ((dateA > dateB) ? 1 : 0));
  });
  
  return playerObjectList;
}

function storeTournyDataInCache () {
  var sheet, cache, tournyDataRange, tournyNames, tournyObject, numOpenTournies, i;
  
  var TOURNYNAMES_COLUMN = 3;
  var TOURNYLINKS_COLUMN = 4;
  var TOURNYMODE_COLUMN = 7;
  var TOURNYENDREGDATE_COLUMN = 8;
  var TOURNYCOUNTRYLIM_COLUMN = 9;
  var TOURNYRANKRANGE_COLUMN = 10;
  var TOURNYTEAMFORMAT_COLUMN = 11;
  var TOURNYTYPE_COLUMN = 15;
  var TOURNYSCORING_COLUMN = 16;
  var STAFFLOOKING_COLUMN = 17;
  var STARTING_ROW = 3;
  
  sheet = SpreadsheetApp.getActive().getSheetByName('Tournaments');
  cache = CacheService.getScriptCache();
  numOpenTournies = getNumOpenTournies(sheet.getRange(STARTING_ROW, TOURNYENDREGDATE_COLUMN, sheet.getLastRow()).getValues()); // Feeds the function the values of the dates column.
  tournyDataRange = sheet.getRange(STARTING_ROW, 1, numOpenTournies, sheet.getLastColumn());
  tournyNames = sheet.getRange(STARTING_ROW, TOURNYNAMES_COLUMN, numOpenTournies).getValues();
  
  for(i = 0; i < numOpenTournies; i++) { // Stores data in cache as key-value pairs.
    tournyObject = {};
    Logger.log("Caching data for " + tournyNames[i]);
    tournyObject[tournyNames[i] + ' LINK'] = '' + sheet.getRange(i+STARTING_ROW, TOURNYLINKS_COLUMN).getValue();
    tournyObject[tournyNames[i] + ' MODE'] = '' + sheet.getRange(i+STARTING_ROW, TOURNYMODE_COLUMN).getValue();
    tournyObject[tournyNames[i] + ' ENDREGDATE'] = '' + sheet.getRange(i+STARTING_ROW, TOURNYENDREGDATE_COLUMN).getValue();
    tournyObject[tournyNames[i] + ' COUNTRYLIM'] = '' + sheet.getRange(i+STARTING_ROW, TOURNYCOUNTRYLIM_COLUMN).getValue();
    tournyObject[tournyNames[i] + ' RANKRANGE'] = '' + sheet.getRange(i+STARTING_ROW, TOURNYRANKRANGE_COLUMN).getValue();
    tournyObject[tournyNames[i] + ' TEAMFORMAT'] = '' + sheet.getRange(i+STARTING_ROW, TOURNYTEAMFORMAT_COLUMN).getValue();
    tournyObject[tournyNames[i] + ' TYPE'] = '' + sheet.getRange(i+STARTING_ROW, TOURNYTYPE_COLUMN).getValue();
    tournyObject[tournyNames[i] + ' SCORING'] = '' + sheet.getRange(i+STARTING_ROW, TOURNYSCORING_COLUMN).getValue();
    tournyObject[tournyNames[i] + ' STAFFLOOKING'] = '' + sheet.getRange(i+STARTING_ROW, STAFFLOOKING_COLUMN).getValue();

    cache.putAll(tournyObject, 600); // Cache information for 10 minutes.
  }
}

function getHtmlBody (playerMail, playerName, matchingArray, otherTourniesArray) {
  var cache, htmlBody, date, matchingArrayString, otherTourniesArrayString, i;
  
  cache = CacheService.getScriptCache(); // Tourny data has already been stored in cache, so it just needs to fetch those while composing the mail.
  
  matchingArrayString = '';
  if (matchingArray.toString().length > 0) { // Add all tournament details to the string only if there are any tournaments that matches the player.
    for (i = 0; i < matchingArray.length; i++) {

      date = cache.get(matchingArray[i] + ' ENDREGDATE');
      // Check if the date is "always open", and if it is not, then recast the variable as a Date type.
      if (date == ' (always open)') {
        // Note that I write exactly " (always open)" (with the space first) in the tournaments sheet because I want it to sort above the dates inside brackets.
        date = "always open";
      } else {
        date = Utilities.formatDate(new Date(date), 'GMT', 'dd/MM/yyyy');
      }
      
      matchingArrayString += '<p> [' + cache.get(matchingArray[i] + ' MODE') + '] <strong><a href="' + cache.get(matchingArray[i] + ' LINK') + '">' + matchingArray[i] + '</a></strong> (' + cache.get(matchingArray[i] + ' TEAMFORMAT') + ') (';
      if (cache.get(matchingArray[i] + ' RANKRANGE').length > 0) { // Only add the rank limit if there is one, otherwise add "no rank limit" to the string.
        matchingArrayString += cache.get(matchingArray[i] + ' RANKRANGE');
      } else {
        matchingArrayString += "no rank limit";
      }
      matchingArrayString += ') (' + date + ') ';
      if (cache.get(matchingArray[i] + ' STAFFLOOKING') == 'Yes') { // Say whether tourny is still looking for staff or not.
        matchingArrayString += "(looking for staff)" + '</p>';
      } else {
        matchingArrayString += '</p>';
      }
    }
  } else {
    matchingArrayString = '<p><i>' + "None at this moment. Check back later!" + '</i></p>';
  }
  
  otherTourniesArrayString = '';
  if (otherTourniesArray.toString().length > 0) {
    for (i = 0; i < otherTourniesArray.length; i++) {
       
      date = cache.get(otherTourniesArray[i] + ' ENDREGDATE');
      // Check if the date is "always open", and if it is not, then recast the variable as a Date type.
      if (date == ' (always open)') {
        // Note that I write exactly " (always open)" (with the space first) in the tournaments sheet because I want it to sort above the dates inside brackets.
        date = "always open";
      } else {
        date = Utilities.formatDate(new Date(date), 'GMT', 'dd/MM/yyyy');
      }
      
      otherTourniesArrayString += '<p> [' + cache.get(otherTourniesArray[i] + ' MODE') + '] <strong><a href="' + cache.get(otherTourniesArray[i] + ' LINK') + '">' + otherTourniesArray[i] + '</a></strong> (' + cache.get(otherTourniesArray[i] + ' TEAMFORMAT') + ') (';
      if (cache.get(otherTourniesArray[i] + ' RANKRANGE').length > 0) { // Only add the rank limit if there is one, otherwise add "no rank limit" to the string.
        otherTourniesArrayString += cache.get(otherTourniesArray[i] + ' RANKRANGE');
      } else {
        otherTourniesArrayString += "no rank limit";
      }
      otherTourniesArrayString += ') (' + date + ') ';
      if (cache.get(otherTourniesArray[i] + ' STAFFLOOKING') == 'Yes') { // Say whether tourny is still looking for staff or not.
        otherTourniesArrayString += "(looking for staff)" + '</p>';
      } else {
        otherTourniesArrayString += '</p>';
      }
    }
  } else {
    otherTourniesArrayString = '<p><i>' + "None at this moment. Check back later!" + '</i></p>';
  }
  
  htmlBody =
    '<head>' +
      '<style>' +
        'h1 {text-align:center;}' +
        'p {text-align:center;}' +
      '</style>' +
  '</head>' +
  '<body>' +
    '<p><img alt="" src="https://osu.ppy.sh/help/wiki/Mascots/img/pippi.png" style="height:250px; width:250px" /></p>' +
    '<h1> ' + '<a href="' + 'https://docs.google.com/spreadsheets/d/1vbtximQJRxr99NsDvtiUghu3vVQo_JNpdanPme2s7_Y/pubhtml' + '">' + "The Osu! Tournament Watch" + '</a></h1>' +
      '<p><span style="font-weight:bold"> ' + "Open tournaments you are eligible for that matches your preferences" + ' </span></p>' +
        matchingArrayString +
      '<p><br><span style="font-weight:bold"> ' + "Other open tournaments" + ' </span></p>' +
        otherTourniesArrayString +
  '</body>';
  
  return htmlBody;
}

function getSubject (playerName) {
  var subject;

  if(playerName.length > 0) { // People can subscribe to the list without entering their usernames. Need to check if they did so that I don't call them "".
    subject = "osu! Tournament Newsletter for " + playerName;
  } else {
    subject = "osu! Tournament Newsletter";
  }
  return subject;
}

function getNumOpenTournies (dateValues) {
  var i;
  
  for (i = 0; dateValues[i].toString().length > 13; i++); // The values in dateValues that are not dates, have lengths of < 13, whereas dates have lengths higher than that.
  return i;
}

function logRemainingDailyQuota () {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Tournaments');
  var value = sheet.getRange('J28').getValue();
  var cache = CacheService.getScriptCache();
  Logger.log(MailApp.getRemainingDailyQuota());
  
  cache.put('MotW - Map of the Week [Monthly]' + ' RANKRANGE2', value);
  Logger.log(cache.get('MotW - Map of the Week [Monthly]' + ' RANKRANGE2'));
  Logger.log(cache.get('MotW - Map of the Week [Monthly]' + ' RANKRANGE2'));
}
