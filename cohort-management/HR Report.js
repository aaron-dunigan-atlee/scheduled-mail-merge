/**
 * Build reports of missing 360s
 *  - Missing manager 360: participant name and manager name
 *  - Participants missing program goal survey (360#1)
 *  - Participants missing DISC
 */
function buildHrReport(templateId, formUrl)
{
  // templateId = templateId || '1j8WyK2RpGqcaE41ASz_BmAO4QBu-Sl4GSjYDedWkdNg'
  // Get template
  var templateDoc = DocumentApp.openById(templateId)
  // Fill template
  var data = getHrReportData(formUrl)
  var body = templateDoc.getBody()
  // Insert tables AND track how many rows were added.
  var rowsAdded = 0;
  rowsAdded += insertReplacementTable(body, "\\[Missing Program Goal Table\\]", [['Missing Program Goal Survey']].concat(data.missingParticipants));
  rowsAdded += insertReplacementTable(body, "\\[Missing Managers Table\\]", [['Participant', 'Manager']].concat(data.missingManagers));
  rowsAdded += insertReplacementTable(body, "\\[Missing DISC Table\\]", [['Missing DISC']].concat(data.missingDisc));
  templateDoc.saveAndClose()

  if (rowsAdded > 0)
  {
    return templateId;
  } else
  {
    // No rows added, so we won't email this template.
    return null;
  }


}

/**
 * @returns {missingParticipants: [Array of names], missingManagers: [Array of [Participant, Manager]]
 */
function getHrReportData(formUrl)
{
  var settings = getCohortSettings();
  var participantData = getParticipantsData(settings, { getBlanks: true })
  console.log("Attempting to open form at " + formUrl)
  var formId = FormApp.openByUrl(formUrl).getId()
  var responseSheet = SpreadsheetApp.getActive().getSheets().find(function (x)
  {
    var formUrlForSheet = x.getFormUrl()
    if (!formUrlForSheet) return false
    return formId == FormApp.openByUrl(formUrlForSheet).getId()
  })
  if (!responseSheet)
  {
    // This error will be caught by sendEmails()
    throw new Error("HR Report: Can't find response sheet for form at " + formUrl)
  }
  console.log("getHrReportData: Getting HR data from sheet " + responseSheet.getName())
  var responseData = getRowsData2(responseSheet, null, { getBlanks: true })
  var hashedResponses = hashObjectsManyToOne(responseData, 'surveyId');

  // Build a list of participants and managers who have not completed the survey
  var missingParticipants = [], missingManagers = [], missingDisc = []
  participantData.forEach(function (participant)
  {
    // Identify missing DISC assessments: If there is no 'DISC Style' on participant list
    if (!participant.discStyle)
    {
      missingDisc.push([participant.participantName])
    }

    // Get responses about this participant
    var responses = hashedResponses[participant.participantId]
    if (!responses)
    {
      // No response for this participant, so we know both responses are missing.
      missingParticipants.push([participant.participantName]);
      missingManagers.push([participant.participantName, participant.managerName]);
      return // to next participant
    }

    // Identify responses by email address
    var participantResponse = responses.find(function (x) { return x.yourEmailAddress == participant.email })
    if (!participantResponse) missingParticipants.push([participant.participantName]);

    // Look for manager response
    var managerResponse = responses.find(function (x) { return x.yourEmailAddress == participant.managerEmail })
    if (!managerResponse) missingManagers.push([participant.participantName, participant.managerName]);

  })


  return { missingParticipants: missingParticipants, missingDisc: missingDisc, missingManagers: missingManagers }
}

/**
 * Find the text, remove it, and get the index within the body, so we can replace with a table.
 * @param {string} text 
 */
function getBodyIndex(body, text)
{
  // body = body || DocumentApp.openById('1j8WyK2RpGqcaE41ASz_BmAO4QBu-Sl4GSjYDedWkdNg').getBody()
  // text = text || "[Missing Managers Table]"
  var foundRangeElement = body.findText(text)
  if (!foundRangeElement)
  {
    console.log("Didn't find field " + text + " in template.")
    return null;
  }
  var paragraph = foundRangeElement.getElement().getParent()
  if (paragraph.getType() != DocumentApp.ElementType.PARAGRAPH)
  {
    console.log("Element parent was not a paragraph.")
    return null;
  }
  var index = body.getChildIndex(paragraph);
  paragraph.removeFromParent();
  console.log("Found " + text + " at child index " + index + " and removed that paragraph.")
  return index;
}

/**
 * Search for the text in body and replace it with a table containing data.
 * @param {DocumentApp.Body} body 
 * @param {string} text  A string containing regex. Yuck. 
 * @param {string[][]} data 
 */
function insertReplacementTable(body, text, data)
{
  // If this table is not referenced in the current template, return rowCount of 0.
  var index = getBodyIndex(body, text);
  if (index === null) return 0;

  // Track how many rows are added.
  var rowCount;
  // Only 1 row? It's just the header, so add an N/A row.
  if (data.length == 1)
  {
    rowCount = 0;
    var naRow = [];
    for (var i = 0; i < data[0].length; i++)
    {
      naRow.push('N/A')
    }
    data = data.concat([naRow])
  } else
  {
    rowCount = data.length - 1;
  }

  body.insertTable(index, data);
  return rowCount;

}