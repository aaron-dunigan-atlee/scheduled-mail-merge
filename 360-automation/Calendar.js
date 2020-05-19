//moved to Master Participant list
function getEvents() {
  var calendar = CalendarApp.getCalendarsByName('Coaching Calls')
  var now = new Date();
  var oneHourFromNow = new Date(now.getTime() + (12 * 60 * 60 * 1000)); //(hours * min * sec * ms) //changed to 12 for testing
  var upcomingEvents = CalendarApp.getDefaultCalendar().getEvents(now, oneHourFromNow);
  if (!upcomingEvents.length) return;
  var coachingCalls = upcomingEvents.filter(function(x){return x});//is a certain color or has a certain tag (TBD)
  var declined = [];
  var nonDeclined = [];
  //filter out the guests who declined the event
  upcomingEvents.forEach(function(event){
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
  console.log('Number of sessions starting in the next hour: ' + upcomingEvents.length);
  
  //get all participant emails
  var accSpreadsheet = SpreadsheetApp.getActive();
  var cohortsSheet = accSpreadsheet.getSheetByName('Cohorts');
  var cohorts = getRowsData(cohortsSheet);
  var activeCohorts = cohorts.filter(function(x){return !x.completed}) //
  
  
 
  activeCohorts.forEach(function(cohort){
    var ss = SpreadsheetApp.openByUrl(cohort.participantList);
    var id = ss.getId();
    var participantSheet = ss.getSheetByName('Participant List');
    var settingsArray = SpreadsheetApp.openByUrl(cohort.cohortManagement).getSheetByName('Cohort Settings').getDataRange().getValues();
    var cohortSettings = {}
    settingsArray.forEach(function(row){if (row[0]) cohortSettings[row[0]] = row[1]});
    var participantData = getRowsData(participantSheet,null,{getMetadata:true});
    var colIndex = participantSheet.getRange('1:1').getValues().shift().indexOf('Timestamps')+1
    if(colIndex){ //will be zero (falsy) if header not found
      var filtered = participantData.filter(function(x){return x.resultsSummary}) //only return rows that have a spreadsheet link; otherwise, there's nothing to send
      //consider filtering out when timestamp in the preceding N columns (matching a string in header) is within an hour (in case we try twice to send alerts)
      filtered.forEach(function(x){
        allParticipants[x.email] = {
          participantListFileId:id,
          sheetRow:x.sheetRow,
          resultsFileId:x.resultsSummary,
          cohortSettings: cohortSettings,
          templateData: getResultsEmailTemplateObject(cohortSettings),
          rowsData: x
        }
      })
    }
  })
  console.log('Found '+Object.keys(allParticipants).length+'participants with results available to send.')  
  
  //Filter down to non-declined guests to upcoming events that match a participant
  nonDeclined.forEach(function(guestEmail){
    if (allParticipants[guestEmail]){ 
      sendResultsEmail(guestEmail, allParticipants[guestEmail])
    }
  })
  
  //These are separate because we don't want Spreadsheet service errors to interfere with sending the emails
  //Log the timestamps
  nonDeclined.forEach(function(update){
    //TODO: group by spreadsheet for fewer calls to SpreadsheetService, possibly Sheets API instead for better batching the ranges
    var sheet = SpreadsheetApp.openById(update.participantListFileId).getSheetByName('Participant List');
    //colIndex determined by how many sheets are on the spreadsheet
    var cell = sheet.getRange(update.colIndex,update.sheetRow);
    var oldValue = cell.getValue();
    cell.setValue(oldValue+'\n'+new Date());
  });
  
  
}


//.match(/[-\w]{25,}/)