function send360ResultsSlack(e)
{
  runWithSlackReporting('send360Results', [e])
}

/**
 * Check whether it is time to send 360 results.  Send them to each participant the night before the coaching call.
 */
function send360Results(e)
{
  var settings = getCohortSettings()

  // Find out if we need to send 360 results today, and if so, which 360
  var todays360 = getTomorrows360(settings)
  if (!todays360)
  {
    console.log("No 360 results to send today.")
    return
  }

  slackCacheLog("Sending 360 results for cohort " + settings["Cohort Name"] + " and 360#" + todays360.number)

  // Get participant data
  var participantSpreadsheet = SpreadsheetApp.openByUrl(settings["Participant List"]);
  var participantListFileId = participantSpreadsheet.getId();
  var participantSheet = participantSpreadsheet.getSheetByName(PARTICIPANT_SHEET_NAME)
  var participantData = getParticipantsData(settings)
  var colIndex = participantSheet.getRange('3:3').getValues().shift().indexOf('Timestamps') + 1

  if (colIndex)
  { //will be zero (falsy) if header not found
    var filtered = participantData.filter(function (x)
    {
      // only return rows that have a valid spreadsheet link; otherwise, there's nothing to send
      if (!x.resultsSummary) return false;
      try
      {
        x.resultsFileId = SpreadsheetApp.openByUrl(x.resultsSummary).getId()
        return true;
      }
      catch (e)
      {
        slackCacheWarn("Participant " + x.participantName + " has invalid Results Summary link: " + x.resultsSummary)
        return false;
      }
    })
    slackCacheLog("Of " + participantData.length + " participants, " + filtered.length + " have results summary links.")

    // Build objects with email data
    var allParticipants = {}
    filtered.forEach(function (x)
    {
      console.log("Building participant object for %s", x.email)
      allParticipants[x.email] = {
        participantListFileId: participantListFileId,
        sheetRow: x.sheetRow,
        resultsFileId: x.resultsFileId,
        cohortSettings: settings,
        templateData: getResultsEmailTemplateObject(settings, todays360),
        rowsData: x,
        colIndex: colIndex
      }
    })
  } else
  {
    slackError("Can't find column 'Timestamps' on Participant List", true)
  }


  filtered.forEach(function (participant)
  {
    var guestEmail = participant.email;
    if (allParticipants[guestEmail])
    {
      try
      {
        sendResultsEmail(guestEmail, allParticipants[guestEmail], e)
      } catch (err)
      {
        slackError(err, true, "Error sending 360 results email.")
      }
    }
  })
  slackCacheLog("Sent 360 results to " + filtered.length + " participants.")

  //These are separate because we don't want Spreadsheet service errors to interfere with sending the emails
  //Log the timestamps
  if (colIndex)
  {
    filtered.forEach(function (participant)
    {
      var guestEmail = participant.email;
      var update = allParticipants[guestEmail]
      if (update)
      {
        var cell = participantSheet.getRange(update.sheetRow, update.colIndex);
        var oldValue = cell.getValue();
        cell.setValue(oldValue + '\n' + new Date());
      }
    });
  }

  var emailSheet = SpreadsheetApp.getActive().getSheetByName('Email Flow')
  todays360.resultSent = new Date()
  setRowsData2(emailSheet, [todays360], { firstRowIndex: todays360.sheetRow })

} // send360Results()

/**
 * Create an object with information for filling a template.
 * @return {object[]} [{name:, templateUrl:, recipient:, subject:, templateValues:}]
 */
function getResultsEmailTemplateObject(cohortSettings, emailFlowRow)
{
  var templateUrl = emailFlowRow.resultEmailTemplate //cohortSettings['Results Email Template']
  if (!templateUrl)
  {
    slackError("No template for sending 360 Results")
  }
  console.log('Building template for Results Email with template url ' + templateUrl)
  var template = {
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
  function buildGlobals(documentUrl)
  {
    var document = DocumentApp.openByUrl(documentUrl)
    var fields = getEmptyTemplateObject(document)
    var globals = {}
    for (var field in fields)
    {
      // Participant fields are not global 
      if (/^Participant:/i.test(field))
      {
        continue;
      }

      // 'Settings:' indicates a Cohort Settings field
      if (/^Settings:/i.test(field))
      {
        var property = field.match(/^Settings: *(.*)/i)[1].trim()
        if (!cohortSettings[property]) slackCacheWarn('Template at ' + documentUrl + ' contains unrecognized field "' + field + '".');
        globals[field] = formatIfDate(cohortSettings[property])

      } else
      {
        slackCacheWarn('Template at ' + documentUrl + ' contains unrecognized field "' + field + '".')
      }
    }
    return globals
  } // getResultsEmailTemplateObject.buildGlobals()

} // getResultsEmailTemplateObject()

/**
 * Build an object with details for creating the email for each participant.
 * @param {Object} template   {name:, templateUrl:, recipientType:, globals:, participantsUrl:, fileId:, rowsData:, formUrl:}
 * @param {Object} settings   From the settings page.  keys are NOT normalized
 * @return {Object}           {recipient:, subject:, fieldValues:}
 */
function buildResultsEmail(template, settings, participant)
{
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
  for (var field in templateFields)
  {
    if (/^Participant:/i.test(field))
    {
      var property = field.match(/^Participant: *(.*)$/i)[1].trim()
      var value = participant[normalizeHeader(property)]
      if (!value) slackCacheWarn('Template at ' + template.documentUrl + ' contains unrecognized field "' + field + '".');
      thisEmail.fieldValues[field] = formatIfDate(value)
    } else
    {
      slackCacheWarn('Unrecognized field: ' + field + ' in template at ' + template.templateUrl)
    }
  }


  return thisEmail;

} // buildResultsEmail()

/**
 * Find out if any 360's are scheduled for tomorrow.
 * Note from Aaron: As of 7.17 we are using the email flow, so we're not checking for "day before the coaching call", we're just checking for "Result Date" on the flow template.
 * @param {Object} settings 
 */
function getTomorrows360(settings)
{
  var emailSheet = SpreadsheetApp.getActive().getSheetByName('Email Flow')
  var emailData = getRowsData2(emailSheet, null, { getMetadata: true })
  var today = new Date()
  var tomorrow = new Date()
  tomorrow.setDate(tomorrow.getDate() + 1)
  // Coaching Call #1 Date is on cohort settings
  // var coachingDate1 = settings["Coaching Call #1 Date"]
  // if (coachingDate1 instanceof Date && today < coachingDate1 && coachingDate1 < tomorrow) {
  //   console.log("Coaching call tomorrow for 360 #1")
  //   return 1
  // }
  // #1 and #2 and #3 are on email flow 
  var threeSixty = emailData.find(function (x)
  {
    // var coachingDate = x.coachingCallDate
    return x.recipient == 'Participant' && x.resultEmailTemplate && x.resultDate instanceof Date && !x.resultSent && x.resultDate < today //&& coachingDate instanceof Date && today < coachingDate && coachingDate < tomorrow
  })
  if (threeSixty && threeSixty.number)
  {
    console.log("Coaching call tomorrow for 360 #" + threeSixty.number)
    return threeSixty
  }
  return null
}

function sendResultsEmail(emailAddress, emailData, e)
{
  var cohortSettings = emailData.cohortSettings;
  var cohortFolder = DriveApp.getFolderById(cohortSettings['Cohort Folder ID'])
  // {name:, templateUrl:, recipientType:, globals:,participantsUrl:}
  var template = emailData.templateData

  // Fill global fields (not participant-specific)
  template.fileId = fillTemplate(template.templateUrl, { 'replacements': template.globals }, cohortFolder, template.name, false)
  // Get list of participant-specific emails to send {recipient:, subject:, fieldValues:}
  var email = buildResultsEmail(template, cohortSettings, emailData.rowsData)

  var replacementObject = { 'replacements': email.fieldValues, 'hyperlinkReplacements': email.hyperlinkReplacements }

  // Fill template.
  console.log("Filling template for email:\n" + JSON.stringify(email, null, 2))
  var emailDocId = fillTemplate(email.templateUrl, replacementObject, cohortFolder, email.name, true)

  var htmlBody = convertDocToHtml(emailDocId)
  DriveApp.getFileById(emailDocId).setTrashed(true)
  // Add attachment if present.
  if (emailData.resultsFileId)
  {
    resultsFileUrl = DriveApp.getFileById(emailData.resultsFileId).getUrl()
    htmlBody += '<br><br><a href ="' + resultsFileUrl + '">Attachment: 360 results</a>'
    // Attempt to share the file with the participant.  Participants may not have gmail addresses, which will throw an error--log it to slack and move on.
    shareSilentyFailSilently(emailData.resultsFileId, emailAddress)
  }

  // Add footer with debug info
  if (TEST_MODE) htmlBody += buildEmailFooter(template.templateUrl, e, 'Send 360 Results')

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
