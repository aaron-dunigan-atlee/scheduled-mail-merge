
/**
 * Load a list of basecamp projects to be returned to the sidebar
 */
function loadBasecampProjects() {
  var projects = getBasecampProjects()
    .map(function(project){
      return {
        'name': project.name, 
        'id': project.id
      }
    });
  if (projects.length === 0) console.log("Basecamp API returned zero projects.")
  return projects
}

function loadBasecampDatesSlack(projectId) {
  return runWithSlackReporting('loadBasecampDates', [projectId])
}

/**
 * Called from the sidebar.  Return the three-sixty dates and session dates for the given project
 */
function loadBasecampDates(projectId) {
  projectId = projectId || 16937451
  var dates = {
    threeSixtyDates: {},
    sessionDates: []
  }
  console.log('Getting basecamp dates for project ' + projectId)
  var project = getBasecampProject(projectId)
  var todoset = project.dock.find(function(x){return x.name == 'todoset'})
  if (!todoset) throw new Error("No to-do set found for project " + projectId)
  var todoLists = getBasecampTodoLists(projectId, todoset.id)
  var todoListTitles = todoLists.map(function(x){return x.title})
  console.log("Found these todo lists:\n" + todoListTitles)
  todoListTitles.forEach(function(title) {
    var dateMatch = title.match(/(\d+)\/(\d+)\/(\d+)$/)
    var sessionMatch = title.match(/^Session\s*(\d+)/i)
    if (sessionMatch && dateMatch) {
      var sessionNumber = sessionMatch[1];
      dates.sessionDates.push({
        'sessionNumber': sessionNumber,
        'date': parseDate(dateMatch)
      })
    }
    var coachingMatch = title.match(/^Coaching.*?(\d)/i)
    if (coachingMatch && dateMatch) {
      var coachingNumber = coachingMatch[1]
      dates.threeSixtyDates[coachingNumber] = parseDate(dateMatch)
    }
  })
  console.log("Got these dates from basecamp: " + JSON.stringify(dates, null, 2))
  return JSON.stringify(dates)

  // Private functions
  // -----------------

  function parseDate(dateMatch) {
    var date = new Date()
    // Just keep everything in UTC to avoid timezone offsets.
    date.setUTCHours(0,0,0,0)
    var month = parseInt(dateMatch[1], 10)-1
    var day = parseInt(dateMatch[2], 10)
    var year = parseInt(dateMatch[3], 10)
    if (year < 2000) year += 2000;
    date.setUTCFullYear(year, month, day)
    return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd')
  }
}

