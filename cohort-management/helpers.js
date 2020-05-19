function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
};


function getAccessToken() {
  var token = ScriptApp.getOAuthToken();
  Logger.log(token)
  return token // DriveApp.getFiles()
}

/**
 * @returns {Object} Keys are column A (NOT normalized); values are Column B
 */
function getCohortSettings() {
  // Get the data we'll need to fill the global values in the templates.
  var ss = SpreadsheetApp.getActive()
  var settingsArray = ss.getSheetByName('Cohort Settings').getDataRange().getValues()
  var cohortSettings = {}
  settingsArray.forEach(function(row){if (row[0]) cohortSettings[row[0]] = row[1]})
  return cohortSettings
}

function formatIfDate(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'America/New_York', 'MMM d, yyyy');
  return value;
}

/**
 * Move Drive file to a destination folder and remove it from all other folders.
 * @param {file} file 
 * @param {folder} destinationFolder 
 */
function moveFile(file, destinationFolder) {
  // Get previous parent folders.
  var oldParents = file.getParents();
  // Add file to destination folder.
  destinationFolder.addFile(file);
  // Remove previous parents.
  while (oldParents.hasNext()) {
    var oldParent = oldParents.next();
    // In case the destination folder was already a parent, don't remove it.
    if (oldParent.getId() != destinationFolder.getId()) {
      oldParent.removeFile(file);
    }
  }
}