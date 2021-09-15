function createCohortSlack(config) {
  return runWithSlackReporting('createCohort', [config])
}

/**
 * Housekeeping to be done when a cohort is first created.
 * @param {Object} config             Configuration object.  Should include these fields:
 * - client         {string} Name of client
 * - flowSheetName  {string} Name of sheet with email flow.
 * - cohortName     {string}
 * - sessionDates   {string}        Stringified JSON array of form [{'sessionNumber': sessionNumber, 'date': dateString}, ...], where dateString is yyyy-MM-dd
 * - coachingCallDates  {string[]}  Dates of coaching call/session for each 360, as formatted by date input field (yyyy-MM-dd).  Email send dates are calculated relative to these.
 * *** EXCEPT for one date, which is relative to Session 1 Date
 */


function createCohortManualTest(){
  var config = {
    "client": "Automation Test",
    "flowSheetName": "Flow Leader Lab",
    "cohortName": "Cohort 5.4",
    "participantList": "https://docs.google.com/spreadsheets/d/1-AFlILXu3r5IoLw_JypC8wLACFLylYjAuHkz7EMQDuA/edit#gid=1825471428",
    "coachingCallDates": [
      "2020-07-28",
      "",
      ""
    ],
    "session1Date": "2020-07-28",
    "sessionDates": ""
  }
  createCohort(config)
}

function createCohort(config) {
  var client = config.client
  var cohortName =  config.cohortName
  console.log("Creating cohort '" + cohortName + "' with this configuration:\n" + JSON.stringify(config, null, 2))
  var coachingCallDates = config.coachingCallDates.map(stringToLocalDate) 
  var session1Date = stringToLocalDate(config.session1Date)
  // var sessionDates = JSON.parse(config.sessionDates || '[]')
  var cohortSuffix = ' - ' + client + ' - ' + cohortName
  
  // Row object to be written to the Cohorts sheet
  var cohortData = {'clientName': client, 'cohortName': cohortName}

  // If non-existent, create client folder.  Create cohort subfolder.
  var rootClientsFolder = DriveApp.getFolderById(CLIENTS_FOLDER_ID);
  var clientFolder = getOrCreateFolderByName(rootClientsFolder, client);
  var cohortFolder = getOrCreateFolderByName(clientFolder, cohortName);
  cohortData.cohortFolder = cohortFolder.getUrl();
  console.log("Cohort folder created: " + cohortData.cohortFolder)
  try{
    shareSilentyFailSilently(cohortFolder.getId(), CALENDAR_ACCOUNT, 'writer')
  } catch(err){
    console.error(err)
  }

  // Create Survey Results subfolder
  var resultsFolder = cohortFolder.createFolder('360 Results' + cohortSuffix)
  config.surveyResultsFolderId = resultsFolder.getId()

  // Set up Participant List
  var participantsListFile
  if (config.participantList) {
    try {
      var participantListSpreadsheet = SpreadsheetApp.openByUrl(config.participantList)
      participantsListFile = DriveApp.getFileById(participantListSpreadsheet.getId())
      console.log("Found participants list at " + config.participantList)
      cohortFolder.addFile(participantsListFile)
    } catch(err) {
      slackCacheWarn("Couldn't open the specified participant list: " + config.participantList + '\n' + err.message)
    }
  }
  if (!participantsListFile) {
    // Create copy of Participant list, rename, and add link to master sheet
    participantsListFile = DriveApp.getFileById(MASTER_PARTICIPANTS_LIST_ID).makeCopy(cohortFolder);
    console.log("Participant List created: " + participantsListFile.getUrl())
  }
  participantsListFile.setName('Participant List' + cohortSuffix)
  cohortData.participantList = participantsListFile.getUrl();
  try{
    shareSilentyFailSilently(participantsListFile.getId(), CALENDAR_ACCOUNT,'writer')
  } catch(err){
    console.error(err)
  }

  // Create copy of Cohort Management 
  var cohortManagementFile = DriveApp.getFileById(MASTER_COHORT_MANAGEMENT_ID).makeCopy(cohortFolder);
  cohortManagementFile.setName('Cohort Management' + cohortSuffix)
  cohortData.cohortManagement = cohortManagementFile.getUrl();
  console.log("Cohort Management created: " + cohortData.cohortManagement)

  // Transfer the email flow, session dates, and settings to Cohort Management
  // First, set the Disc Assessment due date to 1 week before the Session 1 Date
  if (session1Date) {
    var discDueDate = session1Date;
    discDueDate.setDate(discDueDate.getDate() - 7)
    config.discAssessmentDueDate = discDueDate
  }

  var cohortManagementSS = SpreadsheetApp.open(cohortManagementFile)
  setEmailFlow(config.flowSheetName, coachingCallDates, session1Date)
  // Session dates no longer used 7.16.10
  // if (sessionDates.length > 0) {
  //   setSessionDates(sessionDates) 
  //   var firstSession = sessionDates.find(function(x){return x.sessionNumber == 1})
  //   if (firstSession) config.session1Date = firstSession.date
  // }
  config.participantList = cohortData.participantList
  config.cohortFolderId = cohortFolder.getId()
  setCohortSettings(config)
  cohortData.emailFlow = config.flowSheetName

  // Write cohort summary back to Cohorts sheet.
  setRowsData(SpreadsheetApp.getActive().getSheetByName('Cohorts'), [cohortData], {writeMethod: 'append'})
  
  // Log results
  console.log('Created new cohort ' + cohortName + '\nCohort folder at ' + cohortData.cohortFolder)

  // Give the user a confirmation and a link to the new cohort folder.
  showCohortConfirmation(cohortName, cohortData.cohortFolder);

  // Private functions
  // -----------------

  /**
   * Translate the chosen flow sheet into specific dates based on the cohort start date, and transfer it to the Cohort Management spreadsheet.
   * @param {string} flowSheetName 
   * @param {Date[]} coachingCallDates  Dates of coaching calls, in order.
   */
  function setEmailFlow(flowSheetName, coachingCallDates, session1Date) {
    var flowSheet = SpreadsheetApp.getActive().getSheetByName(flowSheetName).copyTo(cohortManagementSS).setName('Email Flow');
    var flowData = getRowsData(flowSheet)
    //   .filter(function(x){
    //     // In case there are more 360's in the template than in the current setup, just remove the additional ones
    //     return (typeof x.number === 'number' && x.number <= coachingCallDates.length)
    //   })
    
    flowData.forEach(function(row){
      // Set 'anchor date' (coaching call or session date) for this 360, from which all due dates are measured.
      var threeSixtyNumber = row.number - 1 // -1 for 0-based array indexing
      if (!coachingCallDates[threeSixtyNumber]) return;
      row.coachingCallDate = coachingCallDates[threeSixtyNumber]

      // Change days +/- to dates
      var dueDateFields = ['surveyDueDate','emailStartDate','emailReminder1Date','emailReminder2Date','emailReminder3Date','resultDate']
      dueDateFields.forEach(function(dueDateField){
        if (typeof row[dueDateField] === 'number' && row.coachingCallDate) {
          // Special exception: two dates are set from the Session 1 Date instead of the coaching call date            
          if (dueDateField == 'emailStartDate' && row.number == 1 && (row.recipient == 'Participant' || row.recipient == 'HR (Participant)')) {
            row[dueDateField] = addBusinessDays(session1Date, row[dueDateField])
          } else {
            row[dueDateField] = addBusinessDays(row.coachingCallDate, row[dueDateField])
          }

        }
      })

    }) // flowData.forEach

    // Write it all back to the sheet.
    setRowsData(flowSheet, flowData, {writeMethod: 'clear'})
    console.log("Assigned email flow '" + flowSheetName + "' with coaching call dates " + coachingCallDates)

  } // createCohort.setEmailFlow()

  /**
   * No longer used as of 7.16.20
   * Write the session dates to the Session Dates sheet
   * @param {Object[]} sessionDates Of the form [{'sessionNumber': sessionNumber, 'date': dateString}, ...], where dateString is yyyy-MM-dd
   */
  // function setSessionDates(sessionDates) {
  //   // No longer using as of 7.16.20
  //   return;
  //   var sessionSheet = cohortManagementSS.getSheetByName('Session Dates');
  //   // Make sure there are no missing numbers
  //   var sessionDatesData = [], expectedSessionNumber = 1
  //   for (var i=0; i<sessionDates.length; i++) {
  //     var thisSessionNumber = parseInt(sessionDates[i].sessionNumber, 10)
  //     while (thisSessionNumber > expectedSessionNumber) {
  //       sessionDatesData.push({'sessionNumber': expectedSessionNumber})
  //       expectedSessionNumber++
  //     }
  //     sessionDatesData.push(sessionDates[i])
  //     expectedSessionNumber++
  //   }
    
  //   // Write it all back to the sheet.
  //   setRowsData(sessionSheet, sessionDatesData, {writeMethod: 'clear'})
  //   console.log("Set session dates " + JSON.stringify(sessionDates))
  // } // createCohort.setSessionDates()

  /**
   * Write the settings to the cohort settings sheet
   * @param {Object} config 
   */
  function setCohortSettings(config) {
    console.log("Assigning Cohort Settings")
    var settingsSheet = cohortManagementSS.getSheetByName('Cohort Settings');
    var range = settingsSheet.getDataRange()
    var settingsArray = range.getValues()
    settingsArray.forEach(function(row){
      var value = config[normalizeHeader(row[0])]
      if (value) {
        row[1] = value;
        console.log(row[0] + ' = ' + value)
      }
    })
    
    // Write it all back to the sheet.
    range.setValues(settingsArray)
    
  } // createCohort.setCohortSettings()

} // createCohort()

function showCohortConfirmation(cohortName, folderUrl){
  var template = HtmlService.createTemplateFromFile('dialog-confirm-cohort');
  template.cohortName = cohortName;
  template.cohortFolderUrl = folderUrl;
  var html = template.evaluate().setWidth(300).setHeight(200)
  SpreadsheetApp.getUi().showModalDialog(html, 'Cohort created');

}
