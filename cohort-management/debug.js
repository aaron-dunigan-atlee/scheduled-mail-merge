// TODO: turn this off after testing
var TEST_MODE = false;
var SCRIPT_VERSION = '08.05'

function buildEmailFooter(templateId, e, source) {
  var footer = "<br/><small><p>This email was sent " + 
    (e ? " by the regular automation process." : " as part of a manual QA test.") + '\n' +
    "Script version: " + SCRIPT_VERSION + '\n' +
    "Template: " + templateId + '\n' +
    (source ? "Source process: " + source : '') + 
    '</p></small>';
  return footer;
}