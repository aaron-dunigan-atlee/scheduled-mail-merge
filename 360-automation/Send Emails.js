function sendEmailsSlack(e) {
  runWithSlackReporting('sendEmails', [e])
}

function sendResultsEmail(emailAddress, emailData) {
  var cohortSettings = emailData.cohortSettings;
  var cohortFolder = DriveApp.getFolderById(cohortSettings['Cohort Folder ID'])
  // {name:, templateUrl:, recipientType:, globals:,participantsUrl:}
  var template = emailData.templateData

    // Fill global fields (not participant-specific)
    template.fileId = fillTemplate(template.templateUrl, {'replacements': template.globals}, cohortFolder, template.name, false)
    // Get list of participant-specific emails to send {recipient:, subject:, fieldValues:}
    var email = buildEmail(template, cohortSettings, emailData.rowsData)
  
      var replacementObject = {'replacements': email.fieldValues, 'hyperlinkReplacements': email.hyperlinkReplacements}

      // Fill template.
      console.log("Filling template for email:\n" + JSON.stringify(email,null,2))
      var emailDocId = fillTemplate(email.templateUrl, replacementObject, cohortFolder, email.name, true)
      var htmlBody = convertDocToHtml(emailDocId).replace(/<style[\s\S]*?<\/style>/ig, '')
      DriveApp.getFileById(emailDocId).setTrashed(true)
      // Add attachment if present.
      if (emailData.resultsFileId) {
        resultsFileUrl = DriveApp.getFileById(emailData.resultsFileId).getUrl()
        htmlBody += '<br><br><a href ="' + resultsFileUrl + '">Attachment: 360 results</a>'
        // Attempt to share the file with the participant.  Participants may not have gmail addresses, which will throw an error--log it to slack and move on.
        shareSilentyFailSilently(emailData.resultsFileId, emailAddress)
      }
      // Send email.
      var message = {
        to: emailAddress,
        bcc: EMAIL_BCC,
        subject: email.subject,
        htmlBody: htmlBody
      }
      MailApp.sendEmail(message)
      console.log("Sent message " + email.subject + " to " + email.recipient)

    DriveApp.getFileById(template.fileId).setTrashed(true)

} // sendResultsEmail()

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
 * Create an object with information for filling a template.
 * @return {object[]} [{name:, templateUrl:, recipient:, subject:, templateValues:}]
 */
function getResultsEmailTemplateObject(cohortSettings) {
  var templateUrl = cohortSettings['Results Email Template']
  if (!templateUrl) {
    slackError("No 'Results Email Template' was specified for cohort " + cohortSettings['Cohort Name'])
  }
  console.log('Building template for Results Email with template url ' + templateUrl )
  var template =  {
    templateUrl: templateUrl,
    name: 'Results Email',
    recipientType: 'Participant',
    globals: buildGlobals(templateUrl),
    participantsUrl: cohortSettings['Participant List'],
  }

  return template

  // Private functions
  // -----------------

  
  /**
   * Build an object of global fields for this document
   * @param {string} documentUrl 
   * @param {Object} flowObject  Object from a row of Email Flow sheet
   */
  function buildGlobals(documentUrl) {
    var document = DocumentApp.openByUrl(documentUrl)
    var fields = getEmptyTemplateObject(document)
    var globals = {}
    for (var field in fields) {
      // Participant fields are not global 
      if (/^Participant:/i.test(field) ) {
        continue;
      }
  
      // 'Settings:' indicates a Cohort Settings field
      if (/^Settings:/i.test(field)) {
        var property = field.match(/^Settings: *(.*)/i)[1].trim()
        if (!cohortSettings[property]) slackCacheWarn('Template at ' + documentUrl + ' contains unrecognized field "' + field + '".');
        globals[field] = formatIfDate(cohortSettings[property])

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
function buildEmail(template, settings, participant) {
  console.log('Building emails for template ' + template.fileId + ', based on ' + template.templateUrl)
  var document = DocumentApp.openById(template.fileId)
  var templateFields = getEmptyTemplateObject(document)
  var subject = extractSubject(document)
  var templateUrl = DriveApp.getFileById(template.fileId).getUrl()
  var thisEmail = {
    'recipient': participant.email,
    'fieldValues': {},
    'subject': subject,
    'templateUrl': templateUrl,
    'name': template.name
  }

  // Create object to fill template for this participant.
  for (var field in templateFields) {
    if (/^Participant:/i.test(field)) {
      var property = field.match(/^Participant: *(.*)$/i)[1].trim()
      var value = participant[normalizeHeader(property)]
      if (!value) slackCacheWarn('Template at ' + template.documentUrl + ' contains unrecognized field "' + field + '".');
      thisEmail.fieldValues[field] = formatIfDate(value)
    } else {
      slackCacheWarn('Unrecognized field: ' + field + ' in template at ' + template.templateUrl)
    }
  }

    
  return thisEmail;

} // buildEmails()

function formatIfDate(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'America/New_York', 'MMM d, yyyy');
  return value;
}