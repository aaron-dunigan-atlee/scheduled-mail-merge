/**
 * Return a url to the form, pre-filled with the values given.
 * @param {string} formUrl
 * @param {Object} fields  Object of form {'Question Text':'prefilled value'} 
 */
function getPrefilledFormUrl(formUrl, fields) {
  var form = FormApp.openByUrl(formUrl)
  var formResponse = form.createResponse()
  var addedItems = false;
  for (var questionText in fields) {
    var formItem = getFormItem(form, questionText)
    if (formItem) {
      var itemResponse = formItem.asTextItem().createResponse(fields[questionText]);
      formResponse = formResponse.withItemResponse(itemResponse)
      console.log("Prefilled response " + fields[questionText] + " in question " + questionText)
      addedItems = true
    } else {
      slackCacheWarn('"' + questionText + '" not found in form at ' + formUrl)
    }
  }
  return addedItems ? formResponse.toPrefilledUrl() : formUrl
}

/** 
* From a form object, get the item object whose title is fieldname.
* @return {FormApp.Item}
*/
function getFormItem(form, fieldName, itemType){
  var items;
  if (itemType) {
    items = form.getItems(itemType);
  } else {
    items = form.getItems(); 
  }
  var foundItem = items.find(function(item){
    return (item.getTitle().trim().toLowerCase() == fieldName.trim().toLowerCase())
  });
  return foundItem;
}