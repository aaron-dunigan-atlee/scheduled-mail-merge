function testEmails() {
  sendEmails()
}

/**
 * Make a copy of a template and fill its fields.
 * @param {string} templateUrl 
 * @param {Object} fieldObject  {replacements: {key: value, ...}, hyperlinkReplacements: [{'field':, 'url':, 'text':}, ...], tableReplacement: }
 * @param {DriveApp.Folder} destinationFolder 
 * @param {string} filename 
 * @param {boolean} replaceEmptyFields  Defaults to true.  If false, leave placeholder if there is no value for the field.
 */
function fillTemplate(templateUrl, fieldObject, destinationFolder, filename, replaceEmptyFields) {
  console.log("Filling template based on " + templateUrl)
  console.log("Template data: \n" + JSON.stringify(fieldObject,null,2))
  if (typeof replaceEmptyFields === 'undefined') replaceEmptyFields = true;
  var replacementObject = fieldObject.replacements || {}
  var hyperlinkReplacements = fieldObject.hyperlinkReplacements || []
  var file = DriveApp.getFileById(getIdFromUrl(templateUrl));
  var templateAsFile = file.makeCopy(filename, destinationFolder);
  var templateId = templateAsFile.getId();
  var templateAsDoc = DocumentApp.openById(templateId);

  // Make table replacements
  if (fieldObject.tableReplacements && countKeys(fieldObject.tableReplacements) > 0) {
    // We're assuming there's only one table and it has at most 2 rows: one for headers and one for fields.
    var table = templateAsDoc.getBody().getTables()[0]
    if (table) {
      var rowCount = table.getNumRows()
      var templateRow = table.getRow(rowCount-1)
      var tableRow = templateRow.copy()
      for (var i=0; i<fieldObject.tableReplacements.length; i++) {
        var tableReplacement = fieldObject.tableReplacements[i]
        table.appendTableRow(
          tableRow.copy().replaceText('>', i+'>').asTableRow()
        )
        for (var field in tableReplacement.fieldValues) {
          replacementObject[field+i] = tableReplacement.fieldValues[field]
        }
        tableReplacement.hyperlinkReplacements.forEach(function(x){
          x.field = x.field + '' + i
          hyperlinkReplacements.push(x)
        })
      }
      // Take out the un-numbered top row.
      templateRow.removeFromParent()
      templateAsDoc.saveAndClose()
      templateAsDoc = DocumentApp.openById(templateId);
    } else {
      slackCacheWarn("Table replacements were given, but no tables were found, for " + templateUrl)
    }

  }


  // Before filling standard fields, replace link field with its hyperlink.
  if (hyperlinkReplacements.length > 0) {
    hyperlinkReplacements.forEach(function(hyperlinkReplacement) {
      addSurveyLinks(templateAsDoc, hyperlinkReplacement)
      // Closed, so re-open
      templateAsDoc = DocumentApp.openById(templateId);
    })
  }

  var requests = [];
  // Add requests for fields.  
  // Start with a template of empty strings in case any fields are missing.
  var templateObject = getEmptyTemplateObject(templateAsDoc);

  templateAsDoc.saveAndClose();
  
  for (var prop in templateObject) {
    if(replacementObject[prop] != undefined){
      templateObject[prop] = replacementObject[prop].toString();
    } 
  }
  requests = buildRequests(templateObject, replaceEmptyFields);
  // Batch update all requests

  if (requests.length > 0) {
    // Requires advanced Docs service
    var response = Docs.Documents.batchUpdate({'requests': requests}, templateId);
    console.log(JSON.stringify(response))
  }

  console.log('templateId is ' + templateId + '\ntemplateAsDoc.getId() is ' + templateAsDoc.getId())
  
  return templateAsDoc.getId();
}

/**
 * Fill the fields in a document (without making a copy).
 * @param {DocumentApp.Document} document 
 * @param {Object} replacementObject 
 * @param {boolean} replaceEmptyFields  Defaults to true.  If false, leave placeholder if there is no value for the field.
 */
function fillDocument(document, replacementObject, replaceEmptyFields) {
  if (typeof replaceEmptyFields === 'undefined') replaceEmptyFields = true;
  var documentId = document.getId()
  var requests = [];
  // Add requests for fields.  
  // Start with a template of empty strings in case any fields are missing.
  var templateObject = getEmptyTemplateObject(document);

  for (var prop in templateObject) {
    if(replacementObject[prop] != undefined){
      templateObject[prop] = replacementObject[prop].toString();
    } 
  }
  requests = buildRequests(templateObject, replaceEmptyFields);

  // Batch update all requests; requires advanced Docs service
  var response = Docs.Documents.batchUpdate({'requests': requests}, documentId);
  console.log(JSON.stringify(response))

  document.saveAndClose();
  return documentId
}
  

/**
 * Find names of all fields, written as {{fieldName}}, in the document,
 * and returns an object with each field assigned an empty string. 
 * @param {Document} document 
 */
function getEmptyTemplateObject(document) {
  var body = document.getBody();
  var searchPattern = /<.*?>/g;
  var bodyText = body.getText();
  var patternMatch = bodyText.match(searchPattern)
  if (!patternMatch) return {}
  var matches = patternMatch.map(function(text){
    return text.slice(1, -1);
  });
  var templateObject = {}
  matches.forEach(function(fieldName){
    templateObject[fieldName] = '';
  });
  return templateObject;
}

/**
 * Create a Docs API request to replace text for each field in replacementObject,
 * and append to the requests array.
 * @param {Array} requests 
 * @param {Object} replacementObject 
 */
function buildRequests(replacementObject, replaceEmptyFields) {
  var requests = []
  for (var prop in replacementObject){
    // i.e. if replaceEmptyFields, replace all, but if not, only replace if there is a value for prop.
    if (replaceEmptyFields || replacementObject[prop]) {
      var request = {
        'replaceAllText': {
          'containsText': {'text': "<"+prop+">", 'matchCase': false},
          'replaceText': formatIfDate(replacementObject[prop])
        }
      };
      console.log('Request made to replace ' + prop + ' with ' + replacementObject[prop])
      requests.push(request);
    }
  }
  return requests;
}

/**
 * For an email template document, remove the header portions to leave just the email body, and return the subject line.
 * @param {DriveApp.Document} document 
 */
function extractSubject(document) {
  // document = document || DocumentApp.openById('1ufsQj6zAMqFqgIHm7TbxcNaVDNWdQ8O-2N_C0YoGQuw') // for testing
  var body = document.getBody()
  var found = body.findText('Subject:')
  if (!found) throw new Error("Email template has no subject line.\n"+document.getUrl())
  
  var subjectLine = found.getElement().getText();
  var subjectIndex = body.getChildIndex(found.getElement().getParent())
  for (var i=0; i<=subjectIndex; i++) {
    body.getChild(0).removeFromParent()
  }
  document.saveAndClose()
  return subjectLine.match(/Subject:(.*)/i)[1].trim()
}

/**
 * Add the pre-filled survey link to a template.
 * Adapted from https://stackoverflow.com/a/39944926
 * hyperlinkObject = {'field':, 'url':, 'text':}
 */
function addSurveyLinks(document, hyperlinkObject) {
  var searchText = '<' + hyperlinkObject.field + '>'
  var body = document.getBody();
  // Find URLs
  var link = body.findText(searchText);
  // Loop through
  while (link != null) {
    // Get the link as an object
    var foundLink = link.getElement().asText();
    // Get the positions of start and end
    var start = link.getStartOffset();
    var end = link.getEndOffsetInclusive();
    // Format link
    foundLink.setLinkUrl(start, end, hyperlinkObject.url);
    foundLink.replaceText(searchText, hyperlinkObject.text)
    // Find next
    link = body.findText(searchText, link);
  }
  document.saveAndClose();
}

function countKeys(object) {
  var count = 0;
  for (var key in object) {
    if (object.hasOwnProperty(key)) {
      count++
    }
  }
  return count
}