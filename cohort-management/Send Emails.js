function sendEmailsSlack(e) {
  runWithSlackReporting('sendEmails', [e])
}

function sendEmails(e) {
  var cohortSettings = getCohortSettings()
  var cohortFolder = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents().next()
  // Get list of email templates to send {name:, templateUrl:, recipientType:, globals:,participantsUrl:}
  var todaysTemplates = getTodaysEmailTemplates(cohortSettings)
  if ((todaysTemplates.templates.length) == 0) {
    slackLog('No emails to send today')
    return
  }
  console.log("Found " + todaysTemplates.templates.length + " email templates to send today.")
  todaysTemplates.templates.forEach(function(template){
    var emailCount = 0
    console.log(template)
    // Fill global fields (not participant-specific)
    template.fileId = fillTemplate(template.templateUrl, {'replacements': template.globals}, cohortFolder, template.name, false)
    // Get list of participant-specific emails to send {recipient:, subject:, fieldValues:}
    var pendingEmails = buildEmails(template, cohortSettings)
  
    // Send each email
    pendingEmails.forEach(function(email){
      // Manager emails need a little extra.
      var replacementObject = {'replacements': email.fieldValues, 'hyperlinkReplacements': email.hyperlinkReplacements}
      if (template.recipientType == 'Manager') {
        replacementObject.tableReplacements = email.participantRows
      }

      // Fill template.
      console.log("Filling template for email:\n" + JSON.stringify(email,null,2))
      var emailDocId = fillTemplate(email.templateUrl, replacementObject, cohortFolder, email.name, true)
      var htmlBody = convertDocToHtml(emailDocId)
      DriveApp.getFileById(emailDocId).setTrashed(true)

      // Send email.
      var message = {
        to: email.recipient,
        bcc: EMAIL_BCC,
        subject: email.subject,
        htmlBody: htmlBody
      }
      // Add attachment if present.
      if (email.attachmentId) {
        var attachment = DriveApp.getFileById(email.attachmentId)
        message.attachments = [attachment]
      }
      MailApp.sendEmail(message)
      console.log("Sent message " + email.subject + " to " + email.recipient)
      emailCount++

    }) // pendingEmails.forEach()

    DriveApp.getFileById(template.fileId).setTrashed(true)
    slackLog("Sent " + emailCount + " emails for " + template.name + " using template at " + template.templateUrl)

  }) // todaysTemplates.templates.forEach()

  if (todaysTemplates.cohortFinished) endCohort(cohortSettings['Cohort Name']);

} // getEmails()

/**
 * Convert Google Doc to HTML
 * Thanks to https://stackoverflow.com/a/28503601
 * @param {string} documentId 
 * Not that this requires Drive Scope but does not automatically force its inclusion in authorizations.  
 * Therefore a reference to DriveApp is needed somewhere in the script.
 */
function convertDocToHtml(documentId) {
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+documentId+"&exportFormat=html";
  var param = 
          {
            method      : "get",
            headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
            muteHttpExceptions:true,
          };
  var html = UrlFetchApp.fetch(url,param).getContentText();
  return html;
}

/**
 * Find out which emails need to be sent today.
 * @return {object[]} [{name:, templateUrl:, recipient:, subject:, templateValues:}]
 */
function getTodaysEmailTemplates(cohortSettings) {
  var today = new Date()
  // If any emails are unsent, we'll update this to false.
  var cohortFinished = true;
  
  // Get data from email flow sheet
  var ss = SpreadsheetApp.getActive()
  var emailFlowSheet = ss.getSheetByName(EMAIL_FLOW_SHEET_NAME);
  var emailFlowObjects = getRowsData2(emailFlowSheet, null, {getMetadata: true})
  // We need to get templates which are next to their due date fields, so we need an array of values.
  var emailFlowArray = emailFlowSheet.getDataRange().getValues()
  var headers = emailFlowArray.shift()

  // Build a list of templates and template names for the emails to be sent
  var templates = [];
  emailFlowObjects.forEach(function(object){
    var row = object.array
    // console.log(row)
    for (var j=0; j<row.length; j++) {
      var property = headers[j]
      // Find date fields, except 'Survey Due Date' and 'Session or Coaching Call Date', which are not email send dates.
      if (property != 'Survey Due Date' &&  property != 'Session or Coaching Call Date' && /Date$/i.test(property) && row[j] instanceof Date) {
        var sendDate = row[j];
        // Sent date is immediately to right of send date.
        var sent = row[j+1]
        if (!sent) cohortFinished = false;
        // Since send dates all have times at midnight, and this triggers after midnight, sendDate < today is sufficient to trigger those scheduled for today.
        if (!sent && sendDate < today) {
          var template =  {
            // 'Send' dates are immediately to the right of template links.
            templateUrl: row[j-1],
            name: property.split(/date/i)[0].trim(),
            recipientType: row[headers.indexOf('Recipient')],
            globals: buildGlobals(row[j-1], object),
            participantsUrl: cohortSettings['Participant List'],
            rowsData: object
          }

//          // If it's a 360 result template, we'll need to attach the 360 results. template.attachmentIds will be a map of {participantId: fileId} for the attachments.
//          if (/result/i.test(template.name)) {
//            template.attachmentIds = compileSurveyResults(object.surveyLink)
//          }
//          console.log('This template to be sent today: ' + template.templateUrl)
//          templates.push(template)
//          // Set the 'sent' column timestamp
//          object[normalizeHeader(headers[j+1])] = today
        }
      } else if (property == 'Session or Coaching Call Date' && row[j] instanceof Date) {
        // Share a folder of all individual 360's with the facilitator the day before the coaching call
        var coachingCallDate = row[j]
        var tomorrow = new Date()
        tomorrow.setDate(tomorrow.getDate()+1)
        if (today < coachingCallDate && coachingCallDate < tomorrow) {
          sendFacilitatorEmail(cohortSettings['Facilitator Email'], cohortSettings['Cohort Folder ID'])
        }
      }
    }
  })
  
  if (templates.length > 0)
  setRowsData2(emailFlowSheet, emailFlowObjects)

  templates = templates.concat(getSessionFeedbackTemplates())

  return {'templates': templates, 'cohortFinished': cohortFinished};

  // Private functions
  // -----------------

  function getSessionFeedbackTemplates() {
    var sessionDatesSheet = SpreadsheetApp.getActive().getSheetByName('Session Dates')
    var sessionDates = getRowsData2(sessionDatesSheet,null,{getMetadata: true})
    var feedbackTemplates = []
    sessionDates.forEach(function(session) {
      if ((session instanceof Date) && session.date < today) {
        var template =  {
          // 'Send' dates are immediately to the right of template links.
          templateUrl: cohortSettings['Session Feedback Email Template'],
          name: 'Session Feedback',
          recipientType: 'Feedback',
          participantsUrl: cohortSettings['Participant List'],
          rowsData: session
        }
        template.rowsData.surveyLink = (session.arrayIndex == sessionDates.length - 1) ? cohortSettings['Final Session Feedback Survey'] : cohortSettings['Session Feedback Survey']
        session.feedbackSurveySent = today;
        feedbackTemplates.push(template)
      } else {
        cohortFinished = false
      }
    })
    if (feedbackTemplates.length > 0) {
      setRowsData2(sessionDatesSheet, sessionDates)
    }
    return feedbackTemplates

  } // getTodaysEmailTemplates.getSessionFeedbackTemplates()

  /**
   * Build an object of global fields for this document
   * @param {string} documentUrl 
   * @param {Object} flowObject  Object from a row of Email Flow sheet
   */
  function buildGlobals(documentUrl, flowObject) {
    var document = DocumentApp.openByUrl(documentUrl)
    var fields = getEmptyTemplateObject(document)
    var globals = {}
    for (var field in fields) {
      // Participant fields are not global and survey links are not global because they are prefilled.
      if (/^Participant:/i.test(field) || /360.*Link/i.test(field)) {
        continue;
      }
  
      // 'Settings:' indicates a Cohort Settings field
      if (/^Settings:/i.test(field)) {
        var property = field.match(/^Settings: *(.*)/i)[1].trim()
        if (!cohortSettings[property]) slackCacheWarn('Template at ' + documentUrl + ' contains unrecognized field "' + field + '".');
        globals[field] = formatIfDate(cohortSettings[property])

      } else if (/^360/.test(field)) {
        if (/Due Date/i.test(field)) {
          globals[field] = formatIfDate(flowObject.surveyDueDate)
        } else {
          slackCacheWarn('Template at ' + documentUrl + ' contains unrecognized field "' + field + '".')
        }

      } else {
        slackCacheWarn('Template at ' + documentUrl + ' contains unrecognized field "' + field + '".')
      }
    }
    return globals
  } // getTodaysEmailTemplates.buildGlobals()

} // getTodaysEmailTemplates()

/**
 * Build an object with details for creating the email for each participant.
 * @param {Object} template   {name:, templateUrl:, recipientType:, globals:, participantsUrl:, fileId:, rowsData:, formUrl:}
 * @param {Object} settings   From the settings page.  keys are NOT normalized
 * @return {Object}           {recipient:, subject:, fieldValues:}
 */
function buildEmails(template, settings) {
  console.log('Building emails for template ' + template.fileId + ', based on ' + template.templateUrl)
  template.formUrl = template.rowsData.surveyLink
  var participantSheet = SpreadsheetApp.openByUrl(template.participantsUrl).getSheetByName('Participant List')
  var participantData = getRowsData2(participantSheet), participantsUpdated = false;
  var document = DocumentApp.openById(template.fileId)
  var templateFields = getEmptyTemplateObject(document)
  var subject = extractSubject(document)
  var isReminderEmail = /reminder/i.test(template.name)
  var templateUrl = DriveApp.getFileById(template.fileId).getUrl()

  // Make a list of individual email objects
  var emails = []

  // If it's a participants email, create one per participant
  if (template.recipientType == 'Participant') {
    participantData.forEach(function(participant){
      if (!participant.participantId) {
        participant.participantId = Utilities.getUuid()
        participantsUpdated = true
      }

      // If it's a reminder email, check if the participant sent the form already.
      if (isReminderEmail) {
        if (alreadySubmitted(participant.participantId, participant.email, template.formUrl)) {
          console.log('No reminder needed. '+ participant.participantName + ' already completed form ' + template.formUrl)
          return
        }
      }

      // Otherwise continue
      var thisEmail = {
        'recipient': participant.email,
        'fieldValues': {},
        'subject': subject,
        'templateUrl': templateUrl,
        'name': template.name
      }

      // Create object to fill template for this participant.
      for (var field in templateFields) {
        if (/^360.*Link/i.test(field)) {
          // Get prefilled survey link
          
          var prefillFields = {
            'Survey ID': participant.participantId,
            'Your email address': participant.email
          }
          prefillFields[settings["Participant Name Field"]] = participant.participantName
          thisEmail.hyperlinkReplacements = [{
            'url': getPrefilledFormUrl(template.formUrl, prefillFields),
            'text': 'Click here',
            'field': field
          }]
        } else if (/^Participant:/i.test(field)) {
          var property = field.match(/^Participant: *(.*)$/i)[1].trim()
          var value = participant[normalizeHeader(property)]
          if (!value) slackCacheWarn('Template at ' + template.documentUrl + ' contains unrecognized field "' + field + '".');
          thisEmail.fieldValues[field] = formatIfDate(value)
        } else {
          slackCacheWarn('Unrecognized field: ' + field + ' in template at ' + template.templateUrl)
        }
      }
      emails.push(thisEmail)
    })
  } // if recipientType == participant

  // TODO: If it's a manager email, summarize all participants in one email
  if (template.recipientType == 'Manager') {
    var managerEmailObjects = {}
    participantData.forEach(function(participant){
      if (!participant.participantId) {
        participant.participantId = Utilities.getUuid()
        participantsUpdated = true
      }

      var thisManager = participant.managerEmail
      if (!managerEmailObjects[thisManager]) {
        managerEmailObjects[thisManager] = {
          'recipient': thisManager,
          'fieldValues': {},
          'subject': subject,
          'templateUrl': templateUrl,
          'name': template.name,
          'participantRows': []
        }
      }

      // If it's a reminder email, check if the participant sent the form already.
      if (isReminderEmail) {
        if (alreadySubmitted(participant.participantId, thisManager, template.formUrl)) {
          console.log('No reminder needed. '+ participant.participantName + ' already completed form ' + template.formUrl)
          return
        }
      }

 
      // Create object to fill template for this participant.
      var thisParticipant = {}
      for (var field in templateFields) {
        if (/^360.*Link/i.test(field)) {
          // Get prefilled survey link
          
          var prefillFields = {
            'Survey ID': participant.participantId,
            'Your email address': thisManager
          }
          prefillFields[settings["Participant Name Field"]] = participant.participantName
          thisParticipant.hyperlinkReplacements = [{
            'url': getPrefilledFormUrl(template.formUrl, prefillFields),
            'text': 'Click here',
            'field': field
          }]
        } else if (/^Participant:/i.test(field)) {
          var property = field.match(/^Participant: *(.*)$/i)[1].trim()
          var value = participant[normalizeHeader(property)]
          if (!value) slackCacheWarn('Template at ' + template.documentUrl + ' contains unrecognized field "' + field + '".');
          thisParticipant.fieldValues[field] = formatIfDate(value)
        } else {
          slackCacheWarn('Unrecognized field: ' + field + ' in template at ' + template.templateUrl)
        }
      }
      managerEmailObjects[thisManager].participantRows.push(thisParticipant)
    })
    
    // Some managers may have empty participantRows, if all their participants already filled the survey. If so, remove them.
    for (var manager in managerEmailObjects) {
      if (managerEmailObjects[manager].participantRows.length > 0) {
        emails.push(managerEmailObjects[manager])
      }
    }

    
  }

  // If it's a direct report, ... 
  if (template.recipientType == 'Direct Report') {
    participantData.forEach(function(participant){
      console.log(participant)
      if (!participant.participantId) {
        participant.participantId = Utilities.getUuid()
        participantsUpdated = true
      }

      // Results emails only go to participants
      if (/Result/i.test(template.name)) {
          return;
      }

      var directReports = participant.directReportsEmails ? participant.directReportsEmails.split(/\s*,\s*/g) : []
      directReports.forEach(function(directReportEmail){

        // If it's a reminder email, check if the participant sent the form already.
        if (isReminderEmail) {
          if (alreadySubmitted(participant.participantId, directReportEmail, template.formUrl)) {
            console.log('No reminder needed. '+ directReportEmail + ' already completed form ' + template.formUrl)
            return
          }
        }

        // Otherwise continue
        var thisEmail = {
          'recipient': participant.email,
          'fieldValues': {},
          'subject': subject,
          'templateUrl': templateUrl,
          'name': template.name
        }
  
        // Create object to fill template for this participant.
        for (var field in templateFields) {
          if (/^360.*Link/i.test(field)) {
            // Get prefilled survey link
            var prefillFields = {
              'Survey ID': participant.participantId,
              'Your email address': directReportEmail
            }
            prefillFields[settings["Participant Name Field"]] = participant.participantName
            thisEmail.hyperlinkReplacements = [{
              'url': getPrefilledFormUrl(template.formUrl, prefillFields),
              'text': 'Click here',
              'field': field
            }]
          } else if (/^Participant:/i.test(field)) {
            var property = field.match(/^Participant: *(.*)$/i)[1].trim()
            var value = participant[normalizeHeader(property)]
            if (!value) slackCacheWarn('Template at ' + template.templateUrl + ' contains unrecognized field "' + field + '".');
            thisEmail.fieldValues[field] = formatIfDate(value)
          } else {
            slackCacheWarn('Unrecognized field: ' + field + ' in template at ' + template.templateUrl)
          }
        }
        emails.push(thisEmail)
      }) // each direct report email
    }) // each participant
  } // if recipient == direct report

  // Feedback surveys are a little different.
  if (template.recipientType == 'Feedback') {
    participantData.forEach(function(participant){
      if (!participant.participantId) {
        participant.participantId = Utilities.getUuid()
        participantsUpdated = true
      }

      var thisEmail = {
        'recipient': participant.email,
        'fieldValues': {},
        'subject': subject,
        'templateUrl': templateUrl,
        'name': template.name
      }

      // Create object to fill template for this participant.
      for (var field in templateFields) {
        if (/Session Feedback Survey/i.test(field)) {
          // Get prefilled survey link
          var prefillFields = {
           // 'Email address': participant.email // Emails are collected automatically on this form.
          }
          thisEmail.hyperlinkReplacements = [{
            'url': getPrefilledFormUrl(template.formUrl, prefillFields),
            'text': 'Click here',
            'field': field
          }]
        } else if (/^Participant:/i.test(field)) {
          var property = field.match(/^Participant: *(.*)$/i)[1].trim()
          var value = participant[normalizeHeader(property)]
          if (!value) slackCacheWarn('Template at ' + template.documentUrl + ' contains unrecognized field "' + field + '".');
          thisEmail.fieldValues[field] = formatIfDate(value)
        } else {
          slackCacheWarn('Template at ' + template.documentUrl + ' contains unrecognized field "' + field + '".')
        }
      }
      emails.push(thisEmail)
    })
  } // if recipientType == feedback

  // If it's a results template, add the attachment file id
  if (/result/i.test(template.name) && template.attachmentIds) {
    emails.forEach(function(email){email.attachmentId = template.attachmentIds[email.recipient]})
  }

  // If needed, write to sheet with updated participant data (uuid, for example)
  if (participantsUpdated) setRowsData2(participantSheet, participantData)
  
  console.log("Built " + emails.length + " emails for this template.")
  
  return emails;

  // Private functions
  // -----------------

  /** Check if the survey with that ID was already submitted */
  function alreadySubmitted(surveyId, respondentEmail, formUrl) {
    var form = FormApp.openByUrl(formUrl);
    var formResponses = form.getResponses()
    var formItems = form.getItems()
    var idItem = formItems.find(function(x){return x.getTitle() == 'Survey ID'})
    if (!idItem) slackError("Form doesn't have a 'Survey ID' field: " + formUrl)
    var emailItem = formItems.find(function(x){return x.getTitle() == 'Your email address'})
    if (!emailItem) slackError("Form doesn't have a 'Your email address' field: " + formUrl)
    var participantResponse = formResponses.find(function(response){ 
      return response.getResponseForItem(idItem).getResponse() === surveyId && response.getResponseForItem(emailItem).getResponse() === respondentEmail
    })
    if (participantResponse) {
      return true;
    } else {
      return false;
    }

  } // buildEmails.alreadySubmitted()

} // buildEmails()

/**
 * Share a folder of all individual 360's with the facilitator the day before the coaching call
 */
function sendFacilitatorEmail(emailAddress, cohortFolderId) {
  var htmlMessage = "<p>Facilitator,</p>"
  htmlMessage += '<p><a href="' + DriveApp.getFolderById(cohortFolderId).getUrl() +'"> Click here </a>'
  htmlMessage += " to view the results summaries for tomorrow's coaching call.</p>"
  MailApp.sendEmail({
    to: emailAddress,
    subject: 'Response summaries for tomorrow\'s coaching call',
    htmlBody: htmlMessage
  })
}