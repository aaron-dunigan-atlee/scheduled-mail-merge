function onFormSubmitSlack(e)
{
  runWithSlackReporting('onFormSubmit', [e])
}

/**
 * For each 360, when form is submitted, 
 * - record program goal to Participant List if requested
 * - record manager goal to Participant List if requested
 * - update rollup report on Automation Control Center
 */
function onFormSubmit(e)
{
  console.log(JSON.stringify(e))
  var participantId = e.namedValues['Survey ID'] // e ? e.namedValues['Survey ID'] : '3adf9181-ef18-489a-aea1-d1ce15a41628' // For testing
  // Feedback surveys won't have a survey ID
  if (!participantId) return;

  var settings = getCohortSettings()

  var participantSpreadsheet = SpreadsheetApp.openByUrl(settings['Participant List']);
  var participantSheet = SpreadsheetApp.openByUrl(settings['Participant List']).getSheetByName(PARTICIPANT_SHEET_NAME)
  var participantData = getParticipantsData(settings)

  var participant = participantData.find(function (x) { return x.participantId == participantId })
  console.log("Participant: " + JSON.stringify(participant))
  if (!participant)
  {
    slackError('Survey ID ' + participantId + ' on form submission does not match any Participant ID')
  }

  var participantUpdated = false;

  // Set participant program goal
  var programGoalArray = e ? e.namedValues[settings['Program Goal Field']] : null
  var programGoal = programGoalArray ? programGoalArray[0] : null;
  console.log('program goal: ' + programGoal)
  if (programGoal && settings['Capture Initial Program Goal'])
  {
    participant.originalProgramGoal = programGoal
    participantUpdated = true;
  }

  // Set manager program goal.  Values in namedValues are arrays
  var managerGoalArray = e ? e.namedValues[settings['Manager Goal Field']] : null
  var managerGoal = managerGoalArray ? managerGoalArray[0] : null;
  console.log('manager goal: ' + managerGoal)
  if (managerGoal && settings['Capture Manager Goal'])
  {
    participant.managerProgramGoal = managerGoal // Is this still used? I don't see it on the spreadsheet. ARDA 9.15.21
    participantUpdated = true;
  }

  // Determine which 360 this is.
  var threeSixtyNumber = getThreeSixtyNumber(e.range.getSheet())

  // Update Individual Data on ACC
  var respondentEmail = e ? e.namedValues['Your email address'] : 'kopaljhalani+participant4.5@gmail.com'
  console.log('Respondent email: ' + respondentEmail)
  if (respondentEmail)
  {
    if (threeSixtyNumber)
    {
      var reportingSheet = SpreadsheetApp.openById(AUTOMATION_CONTROL_CENTER_ID).getSheetByName(INDIVIDUAL_DATA_SHEET_NAME);
      var individualData = getRowsData2(reportingSheet, null, { getMetadata: true })
      var individualRow = individualData.find(function (row) { return row.participantId == participantId && row.roleEmail == respondentEmail && row.number == threeSixtyNumber })
      if (individualRow)
      {
        individualRow.completedYn = 'Yes'
        if (programGoal) individualRow.programGoal = programGoal
        if (managerGoal) individualRow.programGoal = managerGoal
        var range = setRowsData2(reportingSheet, [individualRow], { firstRowIndex: individualRow.sheetRow, endHeader: 'Program Goal' })
        if (range) console.log('Wrote to Individual Data ' + range.getA1Notation())
      } else
      {
        slackError('Can\'t update Individual Data: Email ' + respondentEmail + ' on form submission does not match any "Role Email" on Individual Data', true)
      }
    } else
    {
      slackError('Can\'t update Individual Data or Rollup Data: Couldn\'t match to a 360 number.', true)
    }
  } else
  {
    slackError('Can\'t update Individual Data: No email field on this survey: ' + e.range.getSheet().getFormUrl(), true)
  }

  // Update Rollup Data on ACC
  if (threeSixtyNumber)
  {
    var rollupSheet = SpreadsheetApp.openById(AUTOMATION_CONTROL_CENTER_ID).getSheetByName(ROLLUP_SHEET_NAME);
    var rollupData = getRowsData2(rollupSheet, null, { getMetadata: true })
    var rollupRow = rollupData.find(function (row) { return row.client == settings["Client"] && row.cohort == settings["Cohort Name"] && row.number == threeSixtyNumber })
    if (rollupRow)
    {
      if (typeof rollupRow.complete == 'number')
      {
        rollupRow.complete += 1;
      } else
      {
        rollupRow.complete = 1;
      }
      var range = setRowsData2(rollupSheet, [rollupRow], { firstRowIndex: rollupRow.sheetRow, startHeader: 'Complete', endHeader: 'Complete' })
      if (range) console.log('Wrote to Rollup Data ' + range.getA1Notation() + " with complete: " + rollupRow.complete)
    } else
    {
      slackError('Can\'t update Rollup Data: Couldn\'t match to a 360 number.', true)
    }
  }

  // Update participant's results summary sheet.
  compileSurveyResults()

  // If needed, write back to participant sheet.
  if (participantUpdated)
  {
    var range = setRowsData2(
      participantSpreadsheet.getSheetByName(PROGRAM_GOALS_SHEET_NAME),
      [participant],
      {
        headersRowIndex: 3,
        firstRowIndex: participant.sheetRow, // Note this is the sheet row from the participants tab, but should be the same on the program goals tab
        preserveArrayFormulas: true,
        startHeader: 'Original Program Goal',
        endHeader: 'Original Program Goal',
      }
    )
    if (range) console.log('Wrote goals back to participant sheet ' + range.getA1Notation())

    var range = setRowsData2(
      participantSpreadsheet.getSheetByName(PARTICIPANT_SHEET_NAME),
      [participant],
      {
        headersRowIndex: 3,
        firstRowIndex: participant.sheetRow,
        preserveArrayFormulas: true,
        startHeader: 'Results Summary',
        endHeader: 'Results Summary',
      }
    )
    if (range) console.log('Wrote results summary link back to participant sheet ' + range.getA1Notation())
  }

  return;

  // Private functions
  // -----------------

  function compileSurveyResults()
  {
    var formUrl = e ? e.range.getSheet().getFormUrl() : 'https://docs.google.com/forms/d/1tHesEj_lxi-VVE6uLkvuFJJFdCL2cvuXUaLWqPwEVaM/edit'
    var form = FormApp.openByUrl(formUrl)
    var idItem = getFormItem(form, 'Survey ID')
    // If no Survey Id item, this is not a 360 survey, so don't continue.
    if (!idItem)
    {
      console.log("no Survey Id item, this is not a 360 survey")
      return;
    }

    var resultsSheet = e ? e.range.getSheet() : SpreadsheetApp.getActive().getSheetByName('360 #1');
    // Don't normalize headers in results, because we need to copy these headers over to the results sheet.
    var participantResults = getRowsData2(resultsSheet, null, { headersCase: 'none' }).filter(function (response)
    {
      return response["Survey ID"] == participant.participantId
    })
    if (!participantResults || participantResults.length === 0)
    {
      console.log("No participant results")
      return
    }
    generateResults(participant, participantResults, resultsSheet)

    console.log("Added results for survey " + form.getTitle() + " to results sheet for " + participant.participantName)

    return

    // Private Functions 
    // -----------------


    /**
     * Add a sheet to the participant's results spreadsheet with the current survey's results.
     * @param {Object} participant rowsData object from Participants List
     * @param {Object} participantResults Array of rowsData objects from the survey response sheet.
     */
    function generateResults(participant, participantResults, resultsSheet)
    {
      // Get or create participant's spreadsheet
      var participantSpreadsheet
      if (participant.resultsSummary)
      {
        participantSpreadsheet = SpreadsheetApp.openByUrl(participant.resultsSummary)
        console.log('Using existing participant results spreadsheet')
      } else
      {
        // Create a results file for this participant
        participantSpreadsheet = SpreadsheetApp.create(participant.participantName + ' - 360 Results')
        // Move to the cohort folder
        var file = DriveApp.getFileById(participantSpreadsheet.getId())
        var cohortFolder = DriveApp.getFolderById(settings["Survey Results Folder ID"]);
        moveFile(file, cohortFolder)
        participant.resultsSummary = participantSpreadsheet.getUrl()
        participantUpdated = true;
        console.log('Created results spreadsheet for participant ' + participant.participantName)
      }

      // Get or create sheet for summarizing this 360
      var sheetName = form.getTitle()
      var sheet = participantSpreadsheet.getSheetByName(sheetName)
      if (!sheet)
      {
        var templateSheet = SpreadsheetApp.openById('1WOfZv9CFqy4SPRmNLz1nQutozHeaUDi80xLwpEmMDbc').getSheets()[0];
        sheet = templateSheet.copyTo(participantSpreadsheet).setName(form.getTitle())
      }

      var blankSheet = participantSpreadsheet.getSheetByName('Sheet1')
      if (blankSheet) participantSpreadsheet.deleteSheet(blankSheet)

      console.log("Recording survey results in sheet " + sheet.getName())
      console.log("First row of results is " + JSON.stringify(participantResults[0]))

      // Get headers from the form results sheet (on cohort management)
      var headers = resultsSheet.getRange('1:1').getValues()[0].filter(function (x)
      {
        // Filter for questions we want to report.  Don't report respondent's name or email, timestamp, participant id.
        // Don't report the respondent's relationship, because we categorize by those.
        if (!x || /\bemail\b/i.test(x) || /\bname\b/i.test(x) || /^timestamp$/i.test(x) || /\bsurvey id\b/i.test(x) || x === settings["Respondent Role Field"])
        {
          return false
        }
        return true
      })
      console.log("Summarizing results for these headers: " + headers)

      // Clear previous results and we'll re-compile all of them
      sheet.clearContents()


      // Format results in 4 sections: Manager, Direct Report, Self Reported, Unidentified Respondent
      // For each section, report only columns that were included by that respondent type, i.e. there should not be empty columns in any section.
      var responsesByRole = {
        "DIRECT REPORT": [],
        "MANAGER": [],
        "SELF REPORTED": [],
        "UNIDENTIFIED RESPONDENT": []
      }
      // Track which questions are answered by each respondent type.
      var headersByRole = {
        "DIRECT REPORT": {},
        "MANAGER": {},
        "SELF REPORTED": {},
        "UNIDENTIFIED RESPONDENT": {}
      }

      // var namesByRole = {
      //   "DIRECT REPORT": 'Name not given', 
      //   "MANAGER": participant.managerName, 
      //   "SELF REPORTED": participant.participantName, 
      //   "UNIDENTIFIED RESPONDENT": 'Name not given'
      // }

      participantResults.forEach(function (response)
      {
        console.log("Classifying this response: " + JSON.stringify(response))
        var role = classifyRespondent(response);
        console.log("role is " + role)
        response.respondentRole = role;
        response[role + " FEEDBACK"] = '';
        responsesByRole[role].push(response);
        for (var header in response)
        {
          if (response[header])
          {
            headersByRole[role][header] = true;
          }
        }

      })
      console.log("Headers by role: " + JSON.stringify(headersByRole))
      console.log("Responses by role: " + JSON.stringify(responsesByRole))
      var roles = ["MANAGER", "DIRECT REPORT", "SELF REPORTED", "UNIDENTIFIED RESPONDENT"]

      var data = [], headerRows = [];
      // Track the row (of the sheet) we're building, so we can track header indices to color them later.
      var rowIndex = 1;
      roles.forEach(function (role)
      {
        // Don't print the 'unidentified' category if it's empty.
        if (role == "UNIDENTIFIED RESPONDENT" && responsesByRole[role].length == 0) return; // to next role
        // Add header row: filter them from headers to preserve order
        var roleHeaders = [role + " FEEDBACK"].concat(headers.filter(function (x) { return headersByRole[role][x] }))
        data.push(roleHeaders)
        headerRows.push(rowIndex)
        rowIndex++

        // Add data rows
        for (var r = 0; r < responsesByRole[role].length; r++)
        {
          var newRow = []
          for (var c = 0; c < roleHeaders.length; c++)
          {
            newRow.push(responsesByRole[role][r][roleHeaders[c]])
          }
          data.push(newRow)
          rowIndex++;
        }
        // Blank row for readability... we'll square it up later.
        data.push([])
        rowIndex++

      }) // roles.foreach()

      // Square up the array since rows may not be same length
      makeRowsFlush(data)

      // Write the data
      var width = data[0].length
      sheet.clear()
      sheet.getRange(1, 1, data.length, width)
        .setValues(data)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)

      // Color the header rows
      headerColors = ["#b4dde9", "#b4cbe9", "#b4b5e9", "#c27ba0"]
      for (var i = 0; i < headerRows.length; i++)
      {
        sheet.getRange(headerRows[i], 1, 1, width).setBackground(headerColors[i])
      }

      return participantSpreadsheet

    } // compileSurveyResults.generateResults()

    /**
     * Determine the type of respondent.
     * @param {object} response Rows data object for form response.
     * @returns {string} "DIRECT REPORT", "MANAGER", "SELF REPORTED", or "UNIDENTIFIED RESPONDENT"
     */
    function classifyRespondent(response)
    {
      var roleHeader = settings["Respondent Role Field"]
      if (roleHeader)
      {
        var roleAnswer = response[roleHeader]
        console.log("Respondent's relationship to participant: " + roleAnswer)
        if (roleAnswer)
        {
          if (/direct report/i.test(roleAnswer)) return "DIRECT REPORT";
          if (/manager/i.test(roleAnswer)) return "MANAGER";
          if (/participant/i.test(roleAnswer)) return "SELF REPORTED";
        }
      }
      // Default to this
      console.warn("Unidentified respondent: " + JSON.stringify(response))
      return "UNIDENTIFIED RESPONDENT";

    } // compileSurveyResults.classifyRespondent()

  } // compileSurveyResults()
}

/**
 * Get the 360 number corresponding to a form results sheet.
 * @param {Sheet} sheet 
 */
function getThreeSixtyNumber(sheet)
{
  var url = sheet.getFormUrl()
  if (!url) return null;
  try
  {
    var sheetFormId = FormApp.openByUrl(url).getId();
  } catch (err)
  {
    slackError(err, true)
    return null;
  }

  var emailSheet = SpreadsheetApp.getActive().getSheetByName('Email Flow')
  var emailData = getRowsData2(emailSheet)
  var found = emailData.find(function (x)
  {
    if (!x.surveyLink) return false;
    try
    {
      var formId = FormApp.openByUrl(x.surveyLink).getId()
      return formId == sheetFormId
    } catch (err)
    {
      slackError(err, true, "Couldn't open survey at " + x.surveyLink)
      return false;
    }

  })
  if (found)
  {
    console.log("Sheet " + sheet.getName() + " corresponds to 360#" + found.number)
    return found.number
  } else
  {
    return null;
  }
}


// Make sure all rows are the same length;
function makeRowsFlush(array)
{
  // Find max row length
  var width = 0;
  array.forEach(function (row) { width = Math.max(width, row.length) });
  array.forEach(function (row)
  {
    if (row.length < width)
    {
      for (var i = row.length; i < width; i++)
      {
        row.push('')
      }
    }
  })
}