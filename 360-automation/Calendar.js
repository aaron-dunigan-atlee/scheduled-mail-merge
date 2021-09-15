function getEventsSlack() {
  // runWithSlackReporting('getEvents')
  console.log("No longer using calendar integration.")
}

// No longer used 6.27.20
function getEvents() {
  // var calendar = CalendarApp.getCalendarsByName(COACHING_CALLS_CALENDAR_NAME)
  var settings = getSettings()
  var lookAheadInterval = settings["Calendar Lead Time"];  // Time in hours to look ahead
  if (!(typeof lookAheadInterval === 'number')) {
    slackError("Invalid 'Calendar Lead Time' on ACC Settings sheet.")
  }
  var now = new Date();
  var lookAheadTarget = new Date(now.getTime() + (lookAheadInterval * 60 * 60 * 1000)); //(* min * sec * ms) 
  var lookAheadMinusHour = new Date(lookAheadTarget.getTime() - (1 * 60 * 60 * 1000));  // 1 hour before, to prevent the same event triggering multiple times.  Assumes this function runs on an hourly trigger. 
  // Get the events that *start* up to an hour before the target time.
  var allEvents = CalendarApp.getDefaultCalendar().getEvents(lookAheadMinusHour, lookAheadTarget)
  var upcomingEvents = allEvents.filter(function(x){return lookAheadMinusHour < x.getStartTime() && x.getStartTime() < lookAheadTarget});
  if (!upcomingEvents.length) return;
  var coachingCalls = upcomingEvents.filter(function(x){return x}); //is a certain color or has a certain tag (TBD)
  var declined = [];
  var nonDeclined = [];
  //filter out the guests who declined the event
  coachingCalls.forEach(function(event){
    var guests = event.getGuestList()
    guests.forEach(function(guest){
      var status = guest.getGuestStatus()
      if (status === CalendarApp.GuestStatus.NO){
        declined.push(guest.getEmail());
      } else {
        nonDeclined.push(guest.getEmail());
      }
    })
  })
  console.log('Number of sessions starting in the next ' + lookAheadInterval +' hour(s): ' + coachingCalls.length);
  
  //get all participant emails
  var accSpreadsheet = SpreadsheetApp.getActive();
  var cohortsSheet = accSpreadsheet.getSheetByName('Cohorts');
  var cohorts = getRowsData(cohortsSheet);
  var activeCohorts = cohorts.filter(function(x){return !x.completed}) //
  
  
  var allParticipants = {}
  activeCohorts.forEach(function(cohort){
    var ss = SpreadsheetApp.openByUrl(cohort.participantList);
    var id = ss.getId();
    var participantSheet = ss.getSheetByName('Participant List');
    var settingsArray = SpreadsheetApp.openByUrl(cohort.cohortManagement).getSheetByName('Cohort Settings').getDataRange().getValues();
    var cohortSettings = {}
    settingsArray.forEach(function(row){if (row[0]) cohortSettings[row[0]] = row[1]});
    var participantData = getRowsData(participantSheet,null,{headersRowIndex: 3, getMetadata:true});
    var colIndex = participantSheet.getRange('3:3').getValues().shift().indexOf('Timestamps')+1
    
    if(colIndex){ //will be zero (falsy) if header not found
      var filtered = participantData.filter(function(x){
        // only return rows that have a spreadsheet link; otherwise, there's nothing to send
        return x.resultsSummary
      }) 
      //consider filtering out when timestamp in the preceding N columns (matching a string in header) is within an hour (in case we try twice to send alerts)
      filtered.forEach(function(x){
        allParticipants[x.email] = {
          participantListFileId:id,
          sheetRow: x.sheetRow,
          resultsFileId:SpreadsheetApp.openByUrl(x.resultsSummary).getId(),
          cohortSettings: cohortSettings,
          templateData: getResultsEmailTemplateObject(cohortSettings),
          rowsData: x,
          colIndex: colIndex
        }
      })
    }
  })
  console.log('Found '+Object.keys(allParticipants).length+' participants with results available to send.')  
  if (Object.keys(allParticipants).length == 0) return
  
  //Filter down to non-declined guests to upcoming events that match a participant
  nonDeclined.forEach(function(guestEmail){
    if (allParticipants[guestEmail]){ 
      sendResultsEmail(guestEmail, allParticipants[guestEmail])
    }
  })
  
  //These are separate because we don't want Spreadsheet service errors to interfere with sending the emails
  //Log the timestamps
  nonDeclined.forEach(function(guestEmail){
    var update = allParticipants[guestEmail]
    if (update) {
      //TODO: group by spreadsheet for fewer calls to SpreadsheetService, possibly Sheets API instead for better batching the ranges
      var sheet = SpreadsheetApp.openById(update.participantListFileId).getSheetByName('Participant List');
      //colIndex determined by how many sheets are on the spreadsheet
      var cell = sheet.getRange(update.sheetRow, update.colIndex);
      var oldValue = cell.getValue();
      cell.setValue(oldValue+'\n'+new Date());
    }
  });
  
  
}


//.match(/[-\w]{25,}/)