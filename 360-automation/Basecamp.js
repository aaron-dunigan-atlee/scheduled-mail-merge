
function getBasecampTodoSet(projectId, todoSetId) {
  var todos = makeBasecampRequest('/buckets/' + projectId + '/todosets/' + todoSetId + '.json')
  //console.log(todos)
  return JSON.parse(todos)
}

function getBasecampProjects() {
  var projects = makeBasecampRequest('/projects.json')
  //console.log(projects)
  return JSON.parse(projects)
}

function getBasecampProject(projectId) {
  var project = makeBasecampRequest('/projects/'+ projectId + '.json')
  //console.log(project)
  return JSON.parse(project)
}

function getBasecampTodoLists(projectId, todoSetId) {
  var todos = makeBasecampRequest('/buckets/' + projectId + '/todosets/' + todoSetId + '/todolists.json')
  //console.log(todos)
  return JSON.parse(todos)
}

/**
 * Make a request to the basecamp API.
 * @param {string} endpoint 
 * @param {string} method    Defaults to 'get'
 */
function makeBasecampRequest(endpoint, method) {
  method = method || 'get';
  var url = BASECAMP_API_URL + endpoint;
  console.log('Making Basecamp API request: "' + method + '" ' + url)
  var service = getService();

  var response = UrlFetchApp.fetch(url, {
    'headers': {
      Authorization: 'Bearer ' + service.getAccessToken()
    },
    'method': method
  });
  return response.getContentText()
}

