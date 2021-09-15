function startCohortSlack() {
  runWithSlackReporting('startCohort')
}

/**
 * All the housekeeping for a cohort to go 'live'
 */
function startCohort() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();

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

  // Back up the current file in case anything goes wrong; we'll be able to re-attempt the cohort start.
  var cohortFolder =  DriveApp.getFileById(ss.getId()).getParents().next();
  var backupFolder = cohortFolder.createFolder('Backup while Starting Cohort')
  var backupCohortManagement = DriveApp.getFileById(ss.getId()).makeCopy(ss.getName(),backupFolder)
  var backupProperties = PropertiesService.getScriptProperties().getProperties();
  // Record ID's of existing files.
  var existingFiles = cohortFolder.getFiles()
  var existingFilesIds = []
  while (existingFiles.hasNext()) {
    var file = existingFiles.next()
    existingFilesIds.push(file.getId())
  }
  

  // If any of the following fails, do so gracefully and alert to slack
  var alertMessage;
  try {
    var settings = getCohortSettings();

    // Share the survey results folder with the facilitator
    if (settings['Facilitator 1 Email (Lead)']) {
      try {
        shareSilentyFailSilently(settings['Survey Results Folder ID'], settings['Facilitator 1 Email (Lead)'], 'writer')
        
      } catch(err) {
        slackError(err, true, 'Unable to add facilitator as editor of cohort folder')
      }
    }
    var participantSpreadsheet = SpreadsheetApp.openByUrl(settings['Participant List']);
    var participantSheet = participantSpreadsheet.getSheetByName(PARTICIPANT_SHEET_NAME);
    var participantData = getRowsData2(participantSheet, null, {headersRowIndex: 3});
    
    // Assign UUID's to each participant
    setParticipantIds(participantSheet, participantData);
    
    // Turn on email send trigger
    var hourToSend = settings['Hour To Send'];
    if (!(hourToSend && hourToSend >= 0 && hourToSend <= 23)) hourToSend = 9;
    setEmailSendTrigger(hourToSend);

    // Turn on trigger to send 360 results
    var hourToSend = settings['Hour To Send 360 Results'];
    if (!(hourToSend && hourToSend >= 0 && hourToSend <= 23)) hourToSend = 20;
    setSend360ResultsTrigger(hourToSend)
    
    // Create copy of disc instructions for this cohort
    createDiscInstructions(settings)

    // Copy forms to cohort folder and route responses to this spreadsheet
    // {'rollup': rollupRows, 'individual': individualDataRows}
    var reportingData = setupCohortForms(participantData, settings)
    
    // Add links for facilitators to the Participant List
    addLinksForFacilitators(participantSpreadsheet, settings);

    // Turn on form response trigger
    setFormResponseTrigger()

    // Update the reporting sheets
    addToReporting(reportingData)

    // Record the status
    PropertiesService.getScriptProperties().setProperty('cohort_status', 'Active')
    
    alertMessage = 'This cohort is now active.'
  } catch(err) {
    slackError(err,true,'Error starting cohort ' + SpreadsheetApp.getActive().getUrl())
    try {
      // Delete any new triggers and files and restore the backup.
      var triggers = ScriptApp.getProjectTriggers()
      triggers.forEach(function(trigger){
        try {
          ScriptApp.deleteTrigger(trigger)
        } catch(err) {
          slackError(err,true,"Failed to delete trigger while reverting cohort automation.")
        }
      })
      PropertiesService.getScriptProperties().setProperties(backupProperties)
      var currentSheets = ss.getSheets()
      var backupSs = SpreadsheetApp.open(backupCohortManagement);
      var backupSheetNames = backupSs.getSheets().map(function(x){return x.getName()})
      currentSheets.forEach(function(sheet) {
        var sheetName = sheet.getName()
        if (backupSheetNames.includes(sheetName)) {
          ss.deleteSheet(sheet);
          backupSs.getSheetByName(sheetName).copyTo(ss).setName(sheetName);
        } else {
          // We have to unlink the form before deleting a sheet.
          var formUrl = sheet.getFormUrl();
          if (formUrl) FormApp.openByUrl(formUrl).removeDestination();
          ss.deleteSheet(sheet);
        }
      })
      var allCohortFiles = cohortFolder.getFiles()
      while (allCohortFiles.hasNext()) {
        var file = allCohortFiles.next()
        if (!existingFilesIds.includes(file.getId())) {
          file.setTrashed(true)
        }
      }
      backupCohortManagement.setTrashed(true)
      backupFolder.setTrashed(true)
      alertMessage = 'We\'re sorry, something went wrong and\nwe couldn\'t start this cohort.\n\nCheck that all settings are complete and try again. \n' + err.message

    } catch(err) {
      slackError(err, true, "Error trying to revert cohort after bad start.  Delete this cohort and start from scratch.")
      alertMessage = "Something went REALLY wrong.  You're going to have to create a new cohort and try again. \n" + err.message
    }
  }

  // Give a confirmation
  slackLog("Cohort activated from spreadsheet " + SpreadsheetApp.getActive().getName())
  ui.alert(alertMessage)
  
}

/**
 * Add links for facilitator's reference to the participant list spreadsheet
 * @param {Spreadsheet} spreadsheet Participant spreadsheet
 * @param {Object} settings cohort settings
 * Link to Feedback survey for each session
 * Link to Feedback survey results for each session
 * Link to each 360 survey
 * Link to each of the 360 results folders
 */
function addLinksForFacilitators(spreadsheet, settings) {
  //spreadsheet = spreadsheet || SpreadsheetApp.openById('1hjuk-bQLgFqvy3GL8KHUQkhlQ7tRNXC3wflwAaqnLyE')  // For testing: Cohort 4.8
  settings = settings || getCohortSettings()
  var facilitatorSheet = spreadsheet.getSheetByName('Facilitator Links')
  // Remove any existing sheet and use the most up-to-date template.
  if (facilitatorSheet) {
    spreadsheet.deleteSheet(facilitatorSheet)
  }
  var templateSheet = SpreadsheetApp.openById(FACILITATOR_LINKS_TEMPLATE_ID).getSheetByName('Facilitator Links')
  facilitatorSheet = templateSheet.copyTo(spreadsheet).setName('Facilitator Links')

  var surveyResultsLink = DriveApp.getFileById(settings["Survey Results Folder ID"]).getUrl()

  var sessionCount = parseInt(settings["Number Of Sessions"], 10)
  if (isNaN(sessionCount)) {
    slackError("Can't create facilitator links because Number of Sessions was not provided on Cohort Settings", true);
    return;
  }
  // The template includes a row for "Session 1"... the others need to be inserted.
  if (sessionCount > 1){
    facilitatorSheet.insertRowsAfter(3, sessionCount-1)
    var sessionNames = []
    for (var i=2; i<=sessionCount; i++) {
      sessionNames.push(["Session " + i])
    }
    facilitatorSheet.getRange(4,1, sessionCount - 1, 1).setValues(sessionNames)
    // Just in case, since we need to read the sheet next:
    SpreadsheetApp.flush()
  }

  var data = getRowsData2(facilitatorSheet, null, {headersRowIndex: 2})
  var properties = PropertiesService.getScriptProperties();
  data.forEach(function(row){
    // Session feedback links
    if (row.session === "Session " + sessionCount) {
      row.participantLink = settings["Final Session Feedback Survey"]
      row.results = getFormResultsLink(properties.getProperty("Final Session Feedback Survey"))
    } else if (/^Session \d+$/i.test(row.session)) {
      row.participantLink = settings["Session Feedback Survey"]
      // Properties contains the *edit* url, whereas settings contains the published url.
      row.results = getFormResultsLink(properties.getProperty("Session Feedback Survey"))
    }
    
    // 360 links
    if (/^360 #1/i.test(row.session)) {
      row.participantLink = get360Link(1, 'Direct Report') || get360Link(1, 'Manager')
      row.results = surveyResultsLink
    }
    if (/^360 #2/i.test(row.session)) {
      row.participantLink = get360Link(2, 'Direct Report') || get360Link(2, 'Manager')
      row.results = surveyResultsLink
    }
    if (/^360 #3/i.test(row.session)) {
      row.participantLink = get360Link(3, 'Direct Report') || get360Link(3, 'Manager')
      row.results = surveyResultsLink
    }
    if (/^Final Self/i.test(row.session)) {
      row.participantLink = get360Link(3, 'Participant')
    }
  })
  setRowsData2(facilitatorSheet, data, {headersRowIndex: 2})
}

function get360Link(number, recipient) {
  var emailFlowData = getRowsData2(SpreadsheetApp.getActive().getSheetByName('Email Flow'))
  var this360 = emailFlowData.find(function(x){return x.number == number && x.recipient == recipient})
  if (this360 && this360.surveyLink) {
    return FormApp.openByUrl(this360.surveyLink).getPublishedUrl();
  } else {
    return ''
  } 
}

/**
 * Get a link to the results spreadsheet for a form.
 * @param {string} formUrl 
 */
function getFormResultsLink(formUrl) {
  var resultsId = FormApp.openByUrl(formUrl).getDestinationId()
  return SpreadsheetApp.openById(resultsId).getUrl()
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
  var promptMessage = 'Reminder: this will only work if you filled in all of these sheets.\n(Cells in red need to be filled.)\n\n'
  promptMessage += 'Cohort Settings\nEmail Flow\n\n'
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
function setEmailSendTrigger(hour) {
  // Remove previously set triggers.
  console.log("Setting email send trigger.")
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger){
    if (trigger.getHandlerFunction() === EMAIL_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger)
      console.log("Removed existing trigger.")
    }
  })
  ScriptApp.newTrigger(EMAIL_TRIGGER_FUNCTION)
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .create();
  console.log("Created new trigger.")
}


/**
 * Set a trigger to check daily for sending 360 results.
 */
function setSend360ResultsTrigger(hour) {
  // Remove previously set triggers.
  console.log("Setting 360 results trigger.")
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger){
    if (trigger.getHandlerFunction() === SEND_RESULTS_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger)
      console.log("Removed existing trigger.")
    }
  })
  ScriptApp.newTrigger(SEND_RESULTS_TRIGGER_FUNCTION)
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .create();
  console.log("Created new trigger.")
}

function setupCohortForms(participantData, settings) {
  var ss = SpreadsheetApp.getActive()
  var emailFlowSheet = ss.getSheetByName(EMAIL_FLOW_SHEET_NAME)
  var data = getRowsData2(emailFlowSheet)
  var cohortSuffix = ' - ' + settings['Client'] + ' - ' + settings['Cohort Name']
  var cohortFolder =  DriveApp.getFileById(ss.getId()).getParents().next()
  var formCopies = {} // We will map templateUrl: formCopyUrl
  data.forEach(function(row){
    // Make a copy of the survey.  But the same survey may be linked in more than one row.  In which case, make just one copy.
    var templateFormUrl = row.surveyLink;
    if (!templateFormUrl) return; // 
    if (!formCopies[templateFormUrl]) {
      var formTemplateFile = DriveApp.getFileById(FormApp.openByUrl(templateFormUrl).getId())
      var formCopyName = formTemplateFile.getName() + cohortSuffix
      var formCopyFile = formTemplateFile.makeCopy(formCopyName, cohortFolder)
      var formCopyId = formCopyFile.getId()
      console.log('Made copy ' + formCopyId + ' of original form at ' + templateFormUrl)
      var formCopy = FormApp.openById(formCopyId)
      formCopy.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
      formCopies[templateFormUrl] = formCopy.getEditUrl();
//      send360LinkToFacilitators(formCopy, row.number)  Turned off by client request 6.23.20 Aaron
    }
    // Update the survey link
    row.surveyLink = formCopies[templateFormUrl]
  })
  setRowsData2(emailFlowSheet, data)

  // Do the same for feedback surveys, but put them in a new spreadsheet.
  var properties = PropertiesService.getScriptProperties();
  var feedbackSs = SpreadsheetApp.create('Feedback Survey Results' + cohortSuffix)
  // Move to the cohort folder
  var feedbackSsId = feedbackSs.getId()
  var file = DriveApp.getFileById(feedbackSsId)
  moveFile(file, cohortFolder)
  
  var feedbackSurveys = ['Session Feedback Survey', 'Final Session Feedback Survey']
  feedbackSurveys.forEach(function(field) {
    var templateFormUrl = settings[field]
    if (templateFormUrl) {
      //var formTemplateFile = DriveApp.getFileById(FormApp.openByUrl(templateFormUrl).getId())
      var formTemplateFile = DriveApp.getFileById(getIdFromUrl(templateFormUrl));
      var formCopyName = formTemplateFile.getName() + cohortSuffix
      var formCopyFile = formTemplateFile.makeCopy(formCopyName, cohortFolder)
      var formCopy = FormApp.openById(formCopyFile.getId())
      // Save the edit url in properties for use by the script, without confusing the user.
      properties.setProperty(field, formCopy.getEditUrl());
      // Feedback survey results are sent to participant list.
      formCopy.setDestination(FormApp.DestinationType.SPREADSHEET, feedbackSsId)
      formCopy.setTitle(formCopy.getTitle() + cohortSuffix)
      console.log("Created copy of form for " + field + " called " + formCopyName + " at " + formCopy.getPublishedUrl())
      // Update the settings
      settings[field] = formCopy.getPublishedUrl()
      // And write to the settings sheet
      setCohortSetting(field, settings[field])
    }
  })

  // Delete the blank sheet 
  try {
    var blankSheet = feedbackSs.getSheetByName('Sheet1');
    if (blankSheet) feedbackSs.deleteSheet(blankSheet);
  } catch(err) {
    console.log("Couldn't remove Sheet1 from " + feedbackSs.getUrl() + " because " + err.message)
  }

  // Share the feedback survey results with the facilitator
  shareSilentyFailSilently(feedbackSsId, settings["Facilitator 1 Email (Lead)"])
      
  // Not sure if this is necessary, but want to make sure the form destinations get set before assigning conditional formatting to them.
  SpreadsheetApp.flush()

  // Conditional formatting for form response sheets
  // Color code responses that are “rarely” as red, “sometimes” as yellow and “consistently” as green
  // We set up the rule builders here, but we don't build them yet, because they need to be built and applied to each sheet.
  var ruleBuilders = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Rarely').setBackground("#f4cccc"),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sometimes').setBackground("#fff2cc"),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Often').setBackground("#fff2cc"),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Consistently').setBackground("#b6d7a8")
  ];
  
  // Add the conditional formatting, rename the sheet, AND add the sheet to the 'Rollup' row of Automation Control Center
  var rollupRows = [], individualDataRows= [], spreadsheetUrl = ss.getUrl(), surveysDueCount = getSurveysDueCount();
  console.log('Counted ' + surveysDueCount + ' surveys due for this cohort.')
  ss.getSheets().forEach(addRollupAndFormatting)
  feedbackSs.getSheets().forEach(addRollupAndFormatting)
  return {'rollup': rollupRows, 'individual': individualDataRows}

  // Private functions
  // -----------------

  function addRollupAndFormatting(sheet) {
    var formUrl = sheet.getFormUrl()
    if (formUrl) {
      var range = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns())
      sheet.setConditionalFormatRules(
        ruleBuilders.map(function(rule){
          return rule.setRanges([range]).build();
        })
      );
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
      console.log('Trying to match a flow row to the form with ID ' + sheetFormId)
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
          'complete': '0',
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
        slackCacheWarn("Can't match form response sheet with a 360 Number. Form has ID " + sheetFormId, true)
      }
    }
  } // addRollupAndFormatting()

  function send360LinkToFacilitators(form, threeSixtyNumber) {
    try {
      var facilitator1 = settings['Facilitator 1 Email (Lead)']
      var facilitator2 = settings['Facilitator 2 Email']
      if (!facilitator1) throw new Error("No lead facilitator")
      var facilitators = facilitator1 + (facilitator2 ? ',' + facilitator2 : '')
      var htmlBody = '<p>Dear Facilitator,</p><p>Here is a link to 360 #' + threeSixtyNumber + ' for the upcoming ThinkHuman cohort:</p><p><a href="' + form.getPublishedUrl() + '">360 #' + threeSixtyNumber +'</a></p>'
      htmlBody += '<p>Warmly,</p><p>ThinkHuman</p>'

      MailApp.sendEmail({
        to: facilitators,
        subject: 'Link to 360 #' + threeSixtyNumber,
        htmlBody: htmlBody,
        bcc: EMAIL_BCC
      })
    } catch(err) {
      slackError(err, true, 'Error sending 360 Link to Facilitators')
    }
  }

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
        var direct = Object.assign({}, self)
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
    // Remove previously set triggers.
    console.log("Setting form submit trigger.")
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger){
      if (trigger.getHandlerFunction() === FORM_SUBMIT_TRIGGER_FUNCTION) {
        ScriptApp.deleteTrigger(trigger)
        console.log("Removed existing trigger.")
      }
    })
  ScriptApp.newTrigger(FORM_SUBMIT_TRIGGER_FUNCTION)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();
  console.log("Created new trigger.")
}

/**
 * When cohort is finished:
 * - remove daily email trigger.
 * - set forms to not receive responses.
 *  - mark as completed on ACC
 */
function endCohort(cohortName) {
  slackCacheLog("Shutting down cohort '" + cohortName + "'")
  var ss = SpreadsheetApp.getActive()
  var ssId = ss.getId()
  var triggers = ScriptApp.getProjectTriggers()
  triggers.forEach(function(trigger){
    try {
      ScriptApp.deleteTrigger(trigger)
    } catch(err) {
      slackError(err,true,"Failed to delete trigger while shutting down cohort automation.")
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
  var settings = getCohortSettings()
  var thisCohort = cohortData.find(function(x){
    return x.clientName === settings["Client"] && x.cohortName === settings["Cohort Name"]
  })
  if (thisCohort) {
    thisCohort.completed = new Date();
    setRowsData2(accCohortSheet, [thisCohort], {firstRowIndex: thisCohort.sheetRow, startHeader: 'Completed', endHeader: 'Completed'})
  } else {
    slackError("endCohort: Can't find this cohort on ACC Cohorts sheet to mark it completed.", true)
  }

  // Set status and log to slack
  PropertiesService.getScriptProperties().setProperty('cohort_status','Completed')
  slackCacheLog('Cohort ' + cohortName + ' has completed and all email and form submit triggers have been removed.')
  slackPostCacheLog()
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

function setParticipantIds(participantSheet, participantData) {
  participantData.forEach(function(participant) {
    participant.participantId = participant.participantId || Utilities.getUuid()
  })
  setRowsData2(participantSheet, participantData, {headersRowIndex: 3, startHeader: 'Participant ID', endHeader: 'Participant ID'})
  SpreadsheetApp.flush()
}