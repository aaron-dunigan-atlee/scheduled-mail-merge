/**
 * - Create filled copy of disc instructions. Place in cohort folder.
 *    -- Replace <Participant List Link> with link to PL.
 * - Replace link in cohort settings with link to filled copy.
 * - Set sharing to anyone with link can view.
 * - In Send Emails, replace <Disc Assessment Instructions Link> with link. 
 */


function createDiscInstructions(cohortSettings) {
  var cohortFolder = DriveApp.getFolderById(cohortSettings["Cohort Folder ID"])
  var discInstructionsTemplateUrl = cohortSettings["Disc Instructions Link"]
  if (!discInstructionsTemplateUrl) {
    slackError("Can't generate Disc Instructions: No template provided.", true)
    return
  }

  var replacementObject = buildGlobals(discInstructionsTemplateUrl)

  // Fill template.
  console.log("Filling template for disc instructions")
  var discInstructionsId = fillTemplate(discInstructionsTemplateUrl, replacementObject, cohortFolder, 'DiSC Instructions', true)  
  
  // Set sharing to anyone with link can view
  var file = DriveApp.getFileById(discInstructionsId)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)
  
  // Update cohort settings to link to cohort-specific copy.
  var discInstructionsUrl = DocumentApp.openById(discInstructionsId).getUrl()
  setCohortSetting("Disc Instructions Link", discInstructionsUrl)
  console.log("Disc instructions created: " + discInstructionsUrl)

  return

  // Private functions
  // -----------------

  /**
   * Build an object of global fields for this document
   * @param {string} documentUrl 
   */
  function buildGlobals(documentUrl) {
    var document = DocumentApp.openById(getIdFromUrl(discInstructionsTemplateUrl))
    var fields = getEmptyTemplateObject(document)
    var globals = {}
    for (var field in fields) {  
      // 'Settings:' indicates a Cohort Settings field
      if (/^Settings:/i.test(field)) {
        var property = field.match(/^Settings: *(.*)/i)[1].trim()
        if (!cohortSettings[property]) slackCacheWarn('Template at ' + documentUrl + ' contains unrecognized field "' + field + '".');
        globals[field] = formatIfDate(cohortSettings[property])

      } else if (field == 'Participant List Link') {
        var hyperlinkReplacements = [{
          'url': cohortSettings['Participant List'],
          'text': 'CLICK HERE',
          'field': field
        }]
      } else {
        slackCacheWarn('Template at ' + documentUrl + ' contains unrecognized field "' + field + '".')
      }
    }
    var objectToReturn = {};
    objectToReturn.replacements = globals;
    if (hyperlinkReplacements) objectToReturn.hyperlinkReplacements = hyperlinkReplacements;
    return objectToReturn;
  } // createDiscInstructions.buildGlobals()
}