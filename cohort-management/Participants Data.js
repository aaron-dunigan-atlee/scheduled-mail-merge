/**
 * Get data from the participants list, joining the Participant Information tab with the Program Goals tab
 * @param {Object} settings 
 * @param {Object} options 
 */
function getParticipantsData(settings, options)
{
  // Default options
  var rowsDataOptions = {
    headersRowIndex: 3,
    getMetadata: true
  }
  // Options that are passed to override the defaults
  if (options) Object.assign(rowsDataOptions, options)
  console.log("getParticipantsData: Using these options: %s", rowsDataOptions)
  var participantSpreadsheet = SpreadsheetApp.openByUrl(settings['Participant List']);
  // Participant Information
  var participantSheet = participantSpreadsheet.getSheetByName(PARTICIPANT_SHEET_NAME);
  var participantData = getRowsData2(participantSheet, null, rowsDataOptions);
  // Program Goals
  var goalsSheet = participantSpreadsheet.getSheetByName(PROGRAM_GOALS_SHEET_NAME);
  var goalsData = getRowsData2(goalsSheet, null, rowsDataOptions);
  // console.log(JSON.stringify(goalsData))

  // Join them
  var goalsById = hashObjects(goalsData, 'participantId')
  participantData.forEach(function (participant)
  {
    if (goalsById[participant.participantId])
    {
      participant.originalProgramGoal = goalsById[participant.participantId].originalProgramGoal
      participant.finalProgramGoal = goalsById[participant.participantId].finalProgramGoal
    }
    else
    {
      console.warn("Unable to find participant on Program Goals tab: %s", participant.participantId);
    }
  })
  return participantData
}

function protectProgramGoalColumn(participantSpreadsheet)
{
  var goalsSheet = participantSpreadsheet.getSheetByName(PROGRAM_GOALS_SHEET_NAME);
  var originalProgramGoalColumn = getHeaderColumn(
    goalsSheet,
    'Original Program Goal',
    {
      headersRowIndex: 3
    }
  )
  if (originalProgramGoalColumn === 0)
  {
    throw new Error("Participant list 'Program Goals' tab doesn't have a column for Original Program Goal");
  }
  var goalRange = goalsSheet.getRange(
    4, originalProgramGoalColumn,
    goalsSheet.getMaxRows() - 3, 1
  )
  var protection = goalRange.protect()
  if (protection.canEdit())
  {
    protection.getEditors().forEach(function (editor)
    {
      try
      {
        protection.removeEditor(editor)
      }
      catch (e)
      {
        console.log("Unable to remove editor from program goals column: %s", e.message)
      }
    })
    protection.addEditors([Session.getActiveUser().getEmail(), 'jr@think-human.com'])
      .setDomainEdit(true)
    console.log("Protected range %s on sheet %s", goalRange.getA1Notation(), goalsSheet.getName())
  }
}


/**
 * Find the column index in the header row, for a given header text.
 * @param {Sheet} sheet 
 * @param {string} header Must match exact cell text
 * @param {Object} parameters {'headersRowIndex': integer -- row to look for the headers; defaults to 1}
 * @returns {integer} The sheet column index, or 0 if not found.
 */
function getHeaderColumn(sheet, header, parameters)
{
  parameters = parameters || {}
  var headersIndex = parameters.headersRowIndex || 1;
  var headers = sheet.getRange(headersIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  var column = headers.indexOf(header) + 1
  if (parameters.log !== false) console.log("Header '%s' is on column %s", header, column)
  return column
}