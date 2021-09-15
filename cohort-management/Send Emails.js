function sendEmailsSlack(e) {
  runWithSlackReporting('sendEmails', [e])
}

function sendEmails(e, today) {
  var cohortSettings = getCohortSettings()
  var cohortFolder = DriveApp.getFolderById(cohortSettings["Cohort Folder ID"])
  // Get list of email templates to send {name:, templateUrl:, recipientType:, globals:,participantsUrl:}
  var todaysTemplates = getTodaysEmailTemplates(cohortSettings, today)
  if ((todaysTemplates.templates.length) == 0) {
    slackLog('No emails to send today')
    if (todaysTemplates.cohortFinished) endCohort(cohortSettings['Cohort Name']);
    return
  }
  console.log("Found " + todaysTemplates.templates.length + " email templates to send today.")
  todaysTemplates.templates.forEach(function(template){
    var emailCount = 0, failedEmailCount = 0;
    console.log(template)
    // Fill global fields (not participant-specific)
    template.fileId = fillTemplate(template.templateUrl, {'replacements': template.globals}, cohortFolder, template.name, false)
    // HR emails: Add HR Report tables to the template.
    if (/^HR/.test(template.recipientType || '')) {
      try {
        var hrTemplateId = buildHrReport(template.fileId, template.rowsData.surveyLink)
        if (!hrTemplateId) {
          slackCacheLog("It's time to send the HR report, but there were no missing surveys, so the HR email will not be sent.")
          return; // to next template
        }
      } catch(err) {
        slackError(err, true, "Error building HR Report")
        console.log("Email for this template won't be sent.")
        return; // Don't try to send this template.
      }
    }
    
    // Get list of participant-specific emails to send {recipient:, subject:, fieldValues:}
    var pendingEmails = buildEmails(template, cohortSettings)
    
    // Send each email
    pendingEmails.forEach(function(email){
      try {
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

        // Add footer with debug info
        if (TEST_MODE) htmlBody += buildEmailFooter(template.templateUrl, e, 'Send Emails')

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
      } catch(err) {
        // Continue execution so we can try the next email
        slackError(err, true, "Failed to send email " + email.subject + " to " + email.recipient)
        failedEmailCount++;
      }


    }) // pendingEmails.forEach()

    DriveApp.getFileById(template.fileId).setTrashed(true)
    slackCacheLog("Sent " + emailCount + " emails for " + template.name + " using template at " + template.templateUrl)
    if (failedEmailCount > 0) {
      slackCacheLog("Failed to send " + failedEmailCount + " emails for " + template.name + " using template at " + template.templateUrl)
    }

  }) // todaysTemplates.templates.forEach()

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
  // Remove style tags from the html, except for borders.  Also remove redirects.
  var borderCss = '<style>table {border-collapse: collapse;} table, th, td {border: 1px solid black; padding: 5px;} </style>';
  var html = UrlFetchApp.fetch(url,param).getContentText()
    .replace(/<style[\s\S]*?<\/style>/ig, borderCss)
    .replace(
      /https:\/\/www\.google\.com\/url\?q=(.*?)&.*?\"/g, 
      function(match, captureGroup){return decodeURIComponent(captureGroup) + '"'}
    );
  return html;
}

/**
 * Find out which emails need to be sent today.
 * @return {object[]} [{name:, templateUrl:, recipient:, subject:, templateValues:}]
 */
function getTodaysEmailTemplates(cohortSettings, today) {
  today = today || new Date()
  console.log("Getting email templates for " + today)
  // If any emails are unsent, we'll update this to false.
  var cohortFinished = true;
  var builtFacilitatorEmail = false;
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
      // Find date fields, except 'Survey Due Date' and 'Coaching Call Date', which are not email send dates.
      // And don't send survey results, these go out on a separate trigger.
      if (
        property != 'Survey Due Date' &&  
        property != 'Coaching Call Date' && 
        !(/result/i.test(property)) &&
        /Date$/i.test(property) && 
        row[j] instanceof Date
      ) {
        var sendDate = row[j];
        // Sent date is immediately to right of send date.
        var sent = row[j+1]
        
        var templateUrl = row[j-1]
        // Validate the url
        var isUrl = false
        try {
          if (templateUrl) {
            DocumentApp.openByUrl(templateUrl)
            isUrl = true
          }
        } catch(err) {
          console.log("Not a valid document url: " + templateUrl)
        }
        // Note whether there are still emails unsent.
        if (!sent && templateUrl && isUrl) cohortFinished = false;

        // Since send dates all have times at midnight, and this triggers after midnight, sendDate < today is sufficient to trigger those scheduled for today.
        if (isUrl && !sent && sendDate < today && templateUrl) {
          console.log('Building template for ' + property + ' with template url ' + templateUrl + ' to be sent ' + sendDate + ' to recipient ' + object.recipient + ' for 360#' + object.number)
          var template =  {
            // 'Send' dates are immediately to the right of template links.
            templateUrl: templateUrl,
            name: property.split(/date/i)[0].trim(),
            recipientType: row[headers.indexOf('Recipient')],
            globals: buildGlobals(templateUrl, object),
            participantsUrl: cohortSettings['Participant List'],
            rowsData: object
          }
//          // Results are now sent from a separate trigger
//          // If it's a 360 result template, we'll need to attach the 360 results. template.attachmentIds will be a map of {participantId: fileId} for the attachments.
//          if (/result/i.test(template.name)) {
//            template.attachmentIds = compileSurveyResults(object.surveyLink)
//          }

         console.log('This template to be sent today: ' + JSON.stringify(template))
         templates.push(template)
         // Set the 'sent' column timestamp
         
         object[normalizeHeader(headers[j+1])] = today
        }
      } 

      else if (property == 'Coaching Call Date' && row[j] instanceof Date && !builtFacilitatorEmail) {
        // Share a folder of all individual 360's with the facilitator the day before the coaching call
        // 6.18.20 changed to 2 days before coaching call. Aaron
        
        var coachingCallDate = row[j]
        var tomorrow = TODAY_TEST ? new Date(TODAY_TEST) : new Date()
        tomorrow.setDate(tomorrow.getDate()+1);
        var theNextDay = TODAY_TEST ? new Date(TODAY_TEST) : new Date()
        theNextDay.setDate(theNextDay.getDate()+2);
        if (tomorrow < coachingCallDate && coachingCallDate < theNextDay) {
          console.log("Time to send 360 results to facilitator.")
          builtFacilitatorEmail = true;
          var resultsTemplateUrl = cohortSettings['Facilitator 360 Results Email Template']
          if (resultsTemplateUrl) {
            templates.push(buildFacilitatorEmailTemplate(resultsTemplateUrl))
          }
        } 
      }

    }
  })
  
  if (templates.length > 0)
  setRowsData2(emailFlowSheet, emailFlowObjects)

  // Session feedback surveys no longer sent. 6.27.20
  // templates = templates.concat(getSessionFeedbackTemplates())

  if (cohortFinished) {
    slackCacheLog("It appears that this cohort is finished, so I will turn off automation.")
  }

  return {'templates': templates, 'cohortFinished': cohortFinished};

  // Private functions
  // -----------------

  
  /**
   * Build a template to share a folder of all individual 360's with the facilitator the day before the coaching call
   */
  function buildFacilitatorEmailTemplate(templateUrl) {

    var template =  {
      // 'Send' dates are immediately to the right of template links.
      templateUrl: templateUrl,
      name: '360 Results for Facilitator',
      recipientType: 'Facilitator',
      globals: buildGlobals(templateUrl, {}),
      participantsUrl: cohortSettings['Participant List'],
      rowsData: {}
    }
    console.log("Built template '360 Results for Facilitator' from template " + templateUrl)
    return template
  }

  // No longer used as of 6.27.20.
  // function getSessionFeedbackTemplates() {
  //   var sessionDatesSheet = SpreadsheetApp.getActive().getSheetByName('Session Dates')
  //   var sessionDates = getRowsData2(sessionDatesSheet,null,{getMetadata: true})
  //   var feedbackTemplates = []
  //   var templateUrl = cohortSettings['Session Feedback Email Template'];
  //   // Validate the url
  //   var isUrl = false;
  //   try {
  //     DocumentApp.openByUrl(templateUrl)
  //     isUrl = true
  //   } catch(err) {
  //     console.log("Not a valid document url: " + templateUrl)
  //   }
  //   sessionDates.forEach(function(session) {
  //     if (isUrl && (session.date instanceof Date) && session.date < today) {
  //       var template =  {
  //         templateUrl: templateUrl,
  //         name: 'Session Feedback',
  //         recipientType: 'Feedback',
  //         participantsUrl: cohortSettings['Participant List'],
  //         rowsData: session
  //       }
  //       template.formUrl = (session.arrayIndex == sessionDates.length - 1) ? cohortSettings['Final Session Feedback Survey'] : cohortSettings['Session Feedback Survey']
  //       template.rowsData.surveyLink = template.formUrl
  //       session.feedbackSurveySent = today;
  //       feedbackTemplates.push(template)
  //     } else {
  //       cohortFinished = false
  //     }
  //   })
  //   if (feedbackTemplates.length > 0) {
  //     setRowsData2(sessionDatesSheet, sessionDates)
  //   }
  //   return feedbackTemplates

  // } // getTodaysEmailTemplates.getSessionFeedbackTemplates()

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
  var participantSheet = SpreadsheetApp.openByUrl(template.participantsUrl).getSheetByName(PARTICIPANT_SHEET_NAME)
  var participantData = getRowsData2(participantSheet, null, {headersRowIndex: 3}), participantsUpdated = false;
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
        'name': template.name,
        'hyperlinkReplacements': []
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
          thisEmail.hyperlinkReplacements.push({
            'url': getPrefilledFormUrl(template.formUrl, prefillFields),
            'text': 'CLICK HERE',
            'field': field
          })
        } else if (/^Participant:/i.test(field)) {
          var property = field.match(/^Participant: *(.*)$/i)[1].trim()
          var value = participant[normalizeHeader(property)]
          if (!value) slackCacheWarn('Template at ' + template.templateUrl + ' contains unrecognized or empty field "' + field + '".');
          if (!value && /goal/i.test(property)) value = 'No goal provided.'
          thisEmail.fieldValues[field] = formatIfDate(value)
        } else if (field == 'Disc Assessment Instructions Link') {
          thisEmail.hyperlinkReplacements.push({
            'url': settings['Disc Instructions Link'],
            'text': 'CLICK HERE',
            'field': field
          })
        } else {
          slackCacheWarn('Unrecognized field: ' + field + ' in template at ' + template.templateUrl)
        }
      } // for each participant
      emails.push(thisEmail)
    })
    
  } // if recipientType == participant

  // If it's a manager email, summarize all participants in one email
  if (template.recipientType === 'Manager') {
    var managerEmailObjects = {};
    participantData.forEach(function(participant){
  
//      if (!participant.participantId) {
//        participant.participantId = Utilities.getUuid();
//        participantsUpdated = true;
//      }

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
          console.log('No reminder needed. Manager'+ thisManager + ' already completed form ' + template.formUrl + ' for participant ' + participant.participantName)
          return
        }
      }

 
      // Create object to fill template for this participant.
      var thisParticipant = {'fieldValues':{}}
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
            'text': 'CLICK HERE',
            'field': field
          }]
        } else if (/^Participant:/i.test(field)) {
          var property = field.match(/^Participant: *(.*)$/i)[1].trim()
          var value = participant[normalizeHeader(property)]
          if (!value) slackCacheWarn('Template at ' + template.templateUrl + ' contains unrecognized or empty field "' + field + '".');
          if (!value && /goal/i.test(property)) value = 'No goal provided.'
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
          'recipient': directReportEmail,
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
              'text': 'CLICK HERE',
              'field': field
            }]
          } else if (/^Participant:/i.test(field)) {
            var property = field.match(/^Participant: *(.*)$/i)[1].trim()
            var value = participant[normalizeHeader(property)]
            if (!value) slackCacheWarn('Template at ' + template.templateUrl + ' contains unrecognized or empty field "' + field + '".');
            if (!value && /goal/i.test(property)) value = 'No goal provided.'
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
  // Removed 6.27.20 per client request
  // if (template.recipientType == 'Feedback') {
  //   participantData.forEach(function(participant){
  //     if (!participant.participantId) {
  //       participant.participantId = Utilities.getUuid()
  //       participantsUpdated = true
  //     }

  //     var thisEmail = {
  //       'recipient': participant.email,
  //       'fieldValues': {},
  //       'subject': subject,
  //       'templateUrl': templateUrl,
  //       'name': template.name
  //     }

  //     // Create object to fill template for this participant.
  //     for (var field in templateFields) {
  //       if (/Session Feedback Survey/i.test(field)) {

  //         thisEmail.hyperlinkReplacements = [{
  //           'url': template.formUrl,
  //           'text': 'CLICK HERE',
  //           'field': field
  //         }]
  //       } else if (/^Participant:/i.test(field)) {
  //         var property = field.match(/^Participant: *(.*)$/i)[1].trim()
  //         var value = participant[normalizeHeader(property)]
  //         if (!value) slackCacheWarn('Template at ' + template.templateUrl + ' contains unrecognized field "' + field + '".');
  //         thisEmail.fieldValues[field] = formatIfDate(value)
  //       } else {
  //         slackCacheWarn('Template at ' + template.templateUrl + ' contains unrecognized field "' + field + '".')
  //       }
  //     }
  //     emails.push(thisEmail)
  //   })
    // Removed 6.27.20 per client request. Aaron
    // Add facilitator 1 email
    // var facilitatorEmail = {
    //   'recipient': settings["Facilitator 1 Email (Lead)"],
    //   'fieldValues': {},
    //   'subject': subject,
    //   'templateUrl': templateUrl,
    //   'name': template.name,
    //   'hyperlinkReplacements': [{
    //     'url': template.formUrl,
    //     'text': 'CLICK HERE',
    //     'field': 'Session Feedback Survey'
    //   }]
    // }
    // emails.push(facilitatorEmail)

    // // Add facilitator 2 email
    // var facilitator2Email = {
    //   'recipient': settings["Facilitator 2 Email"],
    //   'fieldValues': {},
    //   'subject': subject,
    //   'templateUrl': templateUrl,
    //   'name': template.name,
    //   'hyperlinkReplacements': [{
    //     'url': template.formUrl, 
    //     'text': 'CLICK HERE',
    //     'field': 'Session Feedback Survey'
    //   }]
    // }
    // emails.push(facilitator2Email)

  // } // if recipientType == feedback

  
  // If it's the facilitator email, create one email
  if (template.recipientType == 'Facilitator') {
    
      // Otherwise continue
      var thisEmail = {
        'recipient': settings['Facilitator 1 Email (Lead)'],
        'fieldValues': {},
        'subject': subject,
        'templateUrl': templateUrl,
        'name': template.name
      }

      // Create object to fill template.
      for (var field in templateFields) {
        if (field === '360 Results Folder' && settings['Survey Results Folder ID']) {
          // Get link to results folder
          var resultsUrl = DriveApp.getFolderById(settings['Survey Results Folder ID']).getUrl()
          thisEmail.hyperlinkReplacements = [{
            'url': resultsUrl,
            'text': 'CLICK HERE',
            'field': field
          }]
        } else {
          slackCacheWarn('Unrecognized field: ' + field + ' in template at ' + template.templateUrl)
        }
      } // for each participant
      emails.push(thisEmail)  
  } // if recipientType = 'Facilitator'

  // If it's the HR report email, create one email
  if (/^HR/.test(template.recipientType || '')) {
  
    var thisEmail = {
      'recipient': settings['Client HR Contact'],
      'fieldValues': {},
      'subject': subject,
      'templateUrl': templateUrl,
      'name': template.name
    }
    emails.push(thisEmail)  
  } // if recipientType = 'HR'

  // If it's a results template, add the attachment file id
  if (/result/i.test(template.name) && template.attachmentIds) {
    emails.forEach(function(email){email.attachmentId = template.attachmentIds[email.recipient]})
  }

  // If needed, write to sheet with updated participant data (uuid, for example)
  if (participantsUpdated) setRowsData2(participantSheet, participantData, {preserveArrayFormulas: true, headersRowIndex: 3})
  
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
