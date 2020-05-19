function onFormSubmitSlack(e) {
  runWithSlackReporting('onFormSubmit',[e])
}

/**
 * For each 360, when form is submitted, 
 * - record program goal to Participant List if requested
 * - record manager goal to Participant List if requested
 * - update rollup report on Automation Control Center
 */
function onFormSubmit(e) {
  var participantId = e.namedValues['Survey ID']
  // Feedback surveys won't have a survey ID
  if (!participantId) return;

  var settings = getCohortSettings()
  
  var participantSheet = SpreadsheetApp.openByUrl(settings['Participant List']).getSheetByName('Participant List')
  var participantData = getRowsData2(participantSheet, null, {getMetadata: true})

  var participant = participantData.find(function(x){return x.participantId == participantId})
  if (!participant) {
    slackError('Survey ID ' + participantId + 'on form submission does not match any Participant ID')
  }

  var participantUpdated = false;

  // Set participant program goal
  var programGoal = e.namedValues[settings['Program Goal Field']]
  console.log('program goal: ')
  if (programGoal && settings['Capture Initial Program Goal']) {
      participant.originalProgramGoal = programGoal
      participantUpdated = true;
  }
  
  // Set manager program goal
  var managerGoal = e.namedValues[settings['Manager Goal Field']]
  console.log('manager goal: ')
  if (managerGoal && settings['Capture Manager Program Goal']) {
      participant.managerProgramGoal = managerGoal
      participantUpdated = true;
  }
  
  // Update Individual Data on ACC
  var respondentEmail = e.namedValues['Your email address']
  if (respondentEmail) {
    var reportingSheet = SpreadsheetApp.openById(AUTOMATION_CONTROL_CENTER_ID).getSheetByName(INDIVIDUAL_DATA_SHEET_NAME);
    var individualData = getRowsData2(reportingSheet, null, {getMetadata: true})
    var individualRow = individualData.find(function(row){ return row.participantId == participantId && row.roleEmail == respondentEmail})
    if (individualRow) {
      individualRow.completedYn = 'Yes'
      if (programGoal) individualRow.programGoal = programGoal
      let range = setRowsData2(reportingSheet, [individualRow], {firstRowIndex: individualRow.sheetRow})
      if (range) console.log('Wrote to Individual Data ' + range.getA1Notation())
    } else {
      slackError('Can\'t update Individual Data: Email ' + respondentEmail + 'on form submission does not match any "Role Email" on Individual Data')
    }
  } else {
    slackError('Can\'t update Individual Data: No email field on this survey: ' + e.range.getSheet().getFormUrl())
  }

  // Update participant's results summary sheet.
  compileSurveyResults()

  // If needed, write back to participant sheet.
  if (participantUpdated) {
    let range = setRowsData2(participantSheet, [participant], {firstRowIndex: participant.sheetRow})
    if (range) console.log('Wrote goals back to participant sheet ' + range.getA1Notation())
  } 

  return;

  // Private functions
  // -----------------

  function compileSurveyResults() {
    var formUrl = e.range.getSheet().getFormUrl()
    var form = FormApp.openByUrl(formUrl)
    var idItem = getFormItem(form, 'Survey ID')
    // If no Survey Id item, this is not a 360 survey, so don't continue.
    if (!idItem) return;

    var formResponses = form.getResponses()
    // Filter for items we'll be reporting the answers to.
    var reportedItems = form.getItems().filter(function(item){
      var itemType = item.getType()
      return (itemType == FormApp.ItemType.LIST || itemType == FormApp.ItemType.GRID)
    })
    
    var participantResults = getReportData(participant)
    if (!participantResults) return 
    generateResults(participant, participantResults)

    console.log("Added results for survey " + form.getTitle() + " to results sheet for " + participant.participantName)
    
    return
  
    // Private Functions 
    // -----------------
  
    /**
     * Summarize all survey responses for this participant
     * @param {Object} participant 
     */
    function getReportData(participant) {
      var id = participant.participantId
      var responsesForParticipant = formResponses.filter(function(response){
        return response.getResponseForItem(item).getResponse() === id;
      })
      if (!responsesForParticipant) {
        slackWarn("Couldn't compile survey results for " + participant.participantName + " because there were no survey responses for that person.")
        return null;
      }
      var reportRows = []
      responsesForParticipant.forEach(function(response){
        var reportRow = {}
        reportedItems.forEach(function(item){
          var itemResponse = response.getResponseForItem(item)
            reportRow[item.getTitle()] = itemResponse.getResponse()
        })
        reportRows.push(reportRow)
      })
    } // compileSurveyResults.generateResultsPdf()
  
    /**
     * Add a sheet to the participant's results spreadsheet with the current survey's results.
     * @param {Object} participant 
     * @param {Object} participantResults 
     */
    function generateResults(participant, participantResults) {
      // Get or create particpant's spreadsheet
      var participantSpreadsheet
      if (participant.resultsSummary) {
        participantSpreadsheet = SpreadsheetApp.openByUrl(participant.resultsSummary)
      } else {
        // Create a results file for this participant
        participantSpreadsheet = SpreadsheetApp.create(participant.participantName + ' - 360 Results')
        // Move to the cohort folder
        var file = DriveApp.getFileById(participantSpreadsheet.getId())
        var cohortFolder = DriveApp.getFolderById(settings["Cohort Folder ID"]);
        moveFile(file, cohortFolder)
        participant.resultsSummary = participantSpreadsheet.getUrl()
        participantsUpdated = true;
        console.log('Created results spreadsheet for participant ' + participant.participantName)
      }

      // Get or create sheet for summarizing this 360
      var sheetName = form.getTitle()
      var sheet = participantSpreadsheet.getSheetByName(sheetName)
      if (!sheet) {
        var templateSheet = SpreadsheetApp.openById('1WOfZv9CFqy4SPRmNLz1nQutozHeaUDi80xLwpEmMDbc').getSheets()[0];
        sheet = templateSheet.copyTo(participantSpreadsheet).setName(form.getTitle())
      }
      
      var headers = []
      for (var prop in participantResults[0]) {
        headers.push(prop)
      }
      var values = []
  
      headers.forEach(function(header){
        var row = [header]
        participantResults.forEach(function(result){
          row.push[result[header]]
        })
        values.push(row)
      })
      
      sheet.getRange(1,1,values.length, values[0].length).setValues(values)
  
      return participantSpreadsheet
  
    } // compileSurveyResults.generateResults()
  
  } // compileSurveyResults()
}
