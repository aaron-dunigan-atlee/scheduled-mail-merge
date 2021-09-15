function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
}

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

function setCohortSetting(field, value) {
  var ss = SpreadsheetApp.getActive()
  var dataRange = ss.getSheetByName('Cohort Settings').getDataRange()
  var settingsArray = dataRange.getValues()
  settingsArray.forEach(function(row){if (row[0] == field) row[1] = value})
  dataRange.setValues(settingsArray)
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


function shareSilentyFailSilently(fileId,userEmail, role){
  role = role || 'reader'
  // Handle email aliases.
  var realEmail = userEmail.replace(/\+.+@/,'@')
  try {
    Drive.Permissions.insert(
    {
      'role': role,
      'type': 'user',
      'value': realEmail
    },
    fileId,
    {
      'sendNotificationEmails': 'false'
    });  
  } catch(err) {
    slackCacheWarn("Couldn't share file " + fileId + " with " + realEmail + ": " + err.message)
  }
}

/**
 * Hash an array of objects by a key
 * @param {Object[]} array 
 * @param {string} key 
 * @param {Object} options
 *    strict {boolean} If true, throw error if key is absent;
 *    keyCase {string} Convert case of key before hashing.  'lower' or 'upper';
 *    verbose {boolean} Log a warning if key is absent;
 * @return {Object} Object of form {key: Object from array}
 */
function hashObjects(array, key, options) {
  options = options || {}
  var hash = {};
  array.forEach(function(object){
    if (object[key]) {
      var thisKey = object[key];
      if (options.keyCase == 'upper') thisKey = thisKey.toLocaleUpperCase();
      if (options.keyCase == 'lower') thisKey = thisKey.toLocaleLowerCase()
      hash[thisKey] = object;
    } else {
      if (options.strict) throw new Error("Can't hash object because it doesn't have key " + key)
      if (options.verbose) console.warn("Can't hash object because it doesn't have key " + key + ": " + JSON.stringify(object))
    }
  })
  return hash
}


/**
 * Hash an array of objects by a key, where there may be multiple elements sharing the same key
 * @param {Object[]} array 
 * @param {string} key 
 * @param {Object} options
 *    strict {boolean} If true, throw error if key is absent;
 *    keyCase {string} Convert case of key before hashing.  'lower' or 'upper';
 *    verbose {boolean} Log a warning if key is absent;
 * @return {Object} Object of form {key: [Objects from array]}
 */
function hashObjectsManyToOne(array, key, options) {
  options = options || {}
  var hash = {};
  array.forEach(function(object){
    if (object[key]) {
      var thisKey = object[key];
      if (options.keyCase == 'upper') thisKey = thisKey.toLocaleUpperCase();
      if (options.keyCase == 'lower') thisKey = thisKey.toLocaleLowerCase()
      if (hash[thisKey]) {
        hash[thisKey].push(object);
      } else {
        hash[thisKey] = [object];
      }
      
    } else {
      if (options.strict) throw new Error("Can't hash object because it doesn't have key " + key)
      if (options.verbose) console.warn("Can't hash object because it doesn't have key " + key + ": " + JSON.stringify(object))
    }
  })
  return hash
}
