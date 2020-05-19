function startCohortSlack() {
  runWithSlackReporting('startCohort')
}

/**
 * All the housekeeping for a cohort to go 'live'
 */
function startCohort() {
  var ui = SpreadsheetApp.getUi();
  // Check if cohort is already active
  var status = PropertiesService.getScriptProperties().getProperty('cohort_status')
  if (status == 'Active') {
    ui.alert('This cohort is already active.')
    return
  }
  // Reminder to fill in settings.
  var continueOk = confirmSettingsFilled()
  if (!continueOk) return;
  // Notify user that emails will come from their account.
  continueOk = confirmEmailSend()
  if (!continueOk) return;
  // If any of the following fails, do so gracefully and alert to slack
  try {
    var settings = getCohortSettings();
    //share the folder
    DriveApp.getFileById(settings['Cohort Folder ID']).addEditor(settings['Facilitator Email']);
    var participantSheet = SpreadsheetApp.openByUrl(settings['Participant List']).getSheetByName('Participant List');
    var participantData = getRowsData2(participantSheet);
    // Turn on email send trigger
    setEmailSendTrigger()
    // Copy forms to cohort folder and route responses to this spreadsheet
    // {'rollup': rollupRows, 'individual': individualDataRows}
    var reportingData = setupCohortForms(participantData, settings)
    // Turn on form response trigger
    setFormResponseTrigger()
    // Update the reporting sheets
    addToReporting(reportingData)
    // Record the status
    PropertiesService.getScriptProperties().setProperty('cohort_status', 'Active')
    // Give a confirmation
    ui.alert('This cohort is now active.')
    slackLog("Cohort activated from spreadsheet " + SpreadsheetApp.getActive().getName())
  } catch(err) {
    slackError(err,true,'Error starting cohort ' + SpreadsheetApp.getActive().getUrl())
    ui.alert('We\'re sorry, something went wrong and\nwe couldn\'t start this cohort.\n\nCheck that all settings are complete and try again.')
  }
}

/**
 * Allow user to confirm that emails will come from their account.
 */
function confirmEmailSend() {
  var ui = SpreadsheetApp.getUi()
  var email;
  try {
    email = Session.getActiveUser().getEmail()
  } catch(err) {
    console.log("Unable to get user email: \n" + err.message)
    email = "Sorry, we couldn't detect your email address."
  }
  var promptMessage = 'Starting the cohort will turn on automatic\nemails at the assigned dates.  Emails\nwill be sent from this account:\n\n' 
  promptMessage += email + '\n\nDo you wish to continue?'
  var response = ui.alert(
    'Start Cohort',
    promptMessage,
    ui.ButtonSet.OK_CANCEL
  )
  if (response == ui.Button.OK) {
    return true
  } else {
    return false
  }
}

/**
 * Reminder to fill in settings, session dates, and email flow.
 */
function confirmSettingsFilled() {
  var ui = SpreadsheetApp.getUi()
  var promptMessage = 'Did you fill in all of these sheets?\n(Cells in red need to be filled.)\n\n'
  promptMessage += 'Cohort Settings\nEmail Flow\nSessionDates\n\n'
  promptMessage += 'Press Cancel to go back and check, OK to continue.'
  var response = ui.alert(
    'Start Cohort',
    promptMessage,
    ui.ButtonSet.OK_CANCEL
  )
  if (response == ui.Button.OK) {
    return true
  } else {
    return false
  }
}

/**
 * Set a trigger to check daily for emails to send.
 */
function setEmailSendTrigger() {
  ScriptApp.newTrigger(EMAIL_TRIGGER_FUNCTION)
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
}

function setupCohortForms(participantData, settings) {
  var ss = SpreadsheetApp.getActive()
  var emailFlowSheet = ss.getSheetByName(EMAIL_FLOW_SHEET_NAME)
  var data = getRowsData2(emailFlowSheet)
  var cohortFolder =  DriveApp.getFileById(ss.getId()).getParents().next()
  var formCopies = {} // We will map templateUrl: formCopyUrl
  data.forEach(function(row){
    // Make a copy of the survey.  But the same survey may be linked in more than one row.  In which case, make just one copy.
    var templateFormUrl = row.surveyLink;
    if (!formCopies[templateFormUrl]) {
      var formCopyFile = DriveApp.getFileById(FormApp.openByUrl(templateFormUrl).getId()).makeCopy(cohortFolder)
      var formCopyId = formCopyFile.getId()
      console.log('Made copy ' + formCopyId + ' of original form at ' + templateFormUrl)
      var formCopy = FormApp.openById(formCopyId)
      formCopy.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
      formCopies[templateFormUrl] = formCopy.getEditUrl();
    }
    // Update the survey link
    row.surveyLink = formCopies[templateFormUrl]
  })
  setRowsData2(emailFlowSheet, data)

  // Do the same for feedback surveys
  var feedbackSurveyUrls = ['Session Feedback Survey', 'Final Session Feedback Survey'].map(function(x){return settings[x]})
  feedbackSurveyUrls.forEach(function(templateFormUrl) {
    if (templateFormUrl) {
      var formCopyFile = DriveApp.getFileById(FormApp.openByUrl(templateFormUrl).getId()).makeCopy(cohortFolder)
      var formCopy = FormApp.openById(formCopyFile.getId())
      formCopy.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
    }
  })

  // Not sure if this is necessary, but want to make sure the form destinations get set before assigning conditional formatting to them.
  SpreadsheetApp.flush()

  // Conditional formatting for form response sheets
  // Color code responses that are “rarely” as red, “sometimes” as yellow and “consistently” as green
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Rarely').setBackground("#f4cccc")
    .whenTextContains('Sometimes').setBackground("#fff2cc")
    .whenTextContains('Often').setBackground("#fff2cc")
    .whenTextContains('Consistently').setBackground("#b6d7a8")
  
  // Add the conditional formatting, rename the sheet, AND add the sheet to the 'Rollup' row of Automation Control Center
  var rollupRows = [], individualDataRows= [], spreadsheetUrl = ss.getUrl(), surveysDueCount = getSurveysDueCount();
  console.log('Counted ' + surveysDueCount + ' surveys due for this cohort.')
  ss.getSheets().forEach(function(sheet){
    var formUrl = sheet.getFormUrl()
    if (formUrl) {
      var range = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns())
      sheet.setConditionalFormatRules([rule.setRanges([range]).build()]);
      var sheetName = FormApp.openByUrl(formUrl).getTitle().slice(0,15);
      
      // Rename sheet according to form title.  make sure no sheet exists with the same name already
      var i = 0;
      do {
        var nameToApply = (i == 0) ? sheetName : sheetName + ' ' + i
        var existingSheet = ss.getSheetByName(nameToApply)
        i++
      } while (existingSheet)
      sheet.setName(nameToApply)

      // Send to rollup
      // Find out which 360 this is.
      var sheetFormId = FormApp.openByUrl(formUrl).getId();
      var surveyNumber, surveyDueDate;
      var matchedFlowRow = data.find(function(x){
        if(x.surveyLink) {
          var formId = FormApp.openByUrl(x.surveyLink).getId()
          console.log('360 #' + x.number + ' has survey id ' + formId)
          return formId === sheetFormId
        } else {
          return false
        }
      })
      if (matchedFlowRow) {
        surveyNumber = matchedFlowRow.number;
        surveyDueDate = matchedFlowRow.surveyDueDate;
        var rollupRow = {
          'complete': buildSubmissionCountFormula(spreadsheetUrl, nameToApply),
          'client': settings['Client'],
          'cohort': settings['Cohort Name'],
          'number': surveyNumber,
          'dueDate': surveyDueDate,
          'total': surveysDueCount,
          'incomplete': '=INDIRECT("R[0]C[1]", FALSE)-INDIRECT("R[0]C[-1]", FALSE)' // Relative R1C1 notation to subtract total - complete
        }
        rollupRows.push(rollupRow)
        addIndividualData(matchedFlowRow, rollupRow)
      } else {
        slackError("Can't match form response sheet with a 360 Number", true)
      }
    }
  })
  return {'rollup': rollupRows, 'individual': individualDataRows}

  // Private functions
  // -----------------

  function buildSubmissionCountFormula(spreadsheetUrl, sheetName) {
    return '=COUNT(IMPORTRANGE("' + spreadsheetUrl + '", "' + sheetName + '!A:A"))'
  } // setupCohortForms.buildSubmissionCountFormula()

  function getSurveysDueCount() {
    var surveyCount = 0
    participantData.forEach(function(participant){
      surveyCount += 2 // one for the participant and one for the manager
      // Multiple direct reports
      if (participant.directReportsEmails) surveyCount += participant.directReportsEmails.split(',').length;
    })
    return surveyCount
  } // setupCohortForms.getSurveysDueCount()


  /**
   * Add to Individual Data: One row for participant, one for manager, one for each direct report
   * @param {Object} flowRow 
   * @param {Object} rollupRow 
   */
  function addIndividualData(flowRow, rollupRow) {
    participantData.forEach(function(participant){
      // One row for participant
      var self = Object.assign({}, rollupRow)
      self.role = 'Participant'
      self.roleEmail = participant.email
      self.participantName = participant.participantName
      self.participantId = participant.participantId
      self.surveyLink = flowRow.surveyLink
      self.completedYn = 'No'
      individualDataRows.push(self)

      // One for manager
      var manager = Object.assign({}, self)
      manager.role = 'Manager'
      manager.roleEmail = participant.managerEmail
      individualDataRows.push(manager)

      // Direct reports
      var directReportEmailsList = (participant.directReportsEmails || '').split(/\s*,\s*/g)
      directReportEmailsList.forEach(function(directReportEmail){
        let direct = Object.assign({}, self)
        direct.role = 'Direct Report'
        direct.roleEmail = directReportEmail
        individualDataRows.push(direct)
      })
      
    })
  
  }  // setupCohortForms.addIndividualData()

} // setupCohortForms()

/**
 * Set a trigger that will copy goals to participant sheet when any form is submitted.
 */
function setFormResponseTrigger() {
  ScriptApp.newTrigger(FORM_SUBMIT_TRIGGER_FUNCTION)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();
}

/**
 * When cohort is finished:
 * - remove daily email trigger.
 * - set forms to not receive responses.
 *  - mark as completed on ACC
 */
function endCohort(cohortName) {
  var ss = SpreadsheetApp.getActive()
  var ssId = ss.getId()
  var triggers = ScriptApp.getProjectTriggers()
  triggers.forEach(function(trigger){
    var handler = trigger.getHandlerFunction()
    if (handler === EMAIL_TRIGGER_FUNCTION || handler === FORM_SUBMIT_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger)
    }
  })
  ss.getSheets().forEach(function(sheet){
    var formUrl = sheet.getFormUrl()
    if (formUrl) {
      FormApp.openByUrl(formUrl).setAcceptingResponses(false)
    }
  })
  
  var accCohortSheet = SpreadsheetApp.openById(AUTOMATION_CONTROL_CENTER_ID).getSheetByName('Cohorts')
  var cohortData = getRowsData2(accCohortSheet,null,{getMetadata: true})
  var thisCohort = cohortData.find(function(x){return SpreadsheetApp.openByUrl(x.cohortManagement).getId() == ssId})
  if (thisCohort) {
    thisCohort.completed = new Date();
    setRowsData2(accCohortSheet, [thisCohort], {firstRowIndex: thisCohort.sheetRow, startHeader: 'Completed', endHeader: 'Completed'})
  } else {
    slackError("endCohort: Can't find this cohort on ACC Cohorts sheet to mark it completed.", true)
  }

  // Set status and log to slack
  PropertiesService.getScriptProperties().setProperty('cohort_status','Completed')
  slackLog('Cohort ' + cohortName + ' has completed and all email and form submit triggers have been removed.')
  
}

/**
 *  Update reporting sheets on ACC
 */
function addToReporting(reportingData) {
  var rollupSheet = SpreadsheetApp.openById(AUTOMATION_CONTROL_CENTER_ID).getSheetByName(ROLLUP_SHEET_NAME);
  var range = setRowsData2(rollupSheet, reportingData.rollup, {writeMethod: 'append'});
  if (range) {
    console.log('Added rows to Rollup on ' + range.getA1Notation())
  }
  var individualSheet = SpreadsheetApp.openById(AUTOMATION_CONTROL_CENTER_ID).getSheetByName(INDIVIDUAL_DATA_SHEET_NAME);
  var range = setRowsData2(individualSheet, reportingData.individual, {writeMethod: 'append'});
  if (range) {
    console.log('Added rows to Individual Data on ' + range.getA1Notation())
  }
}

