

function testPreFilledLinks()
{
  var settings = getCohortSettings()
  var participantData = getParticipantsData(settings)
  participantData.forEach(function (participant)
  {
    // Get prefilled survey link: Survey ID and Participant Name applied to manager indirect as well; email address will be updated for each below
    var prefillFields = {
      'Survey ID': participant.participantId,
      'Your email address': participant.email
    }
    prefillFields[settings["Participant Name Field"]] = participant.participantName
    var surveyLink = getPrefilledFormUrl(
      "https://docs.google.com/forms/d/1pO7dHHswcNBSvfD4ZnHFODvIHgijJ1TILdWbGLgpG94/edit",
      prefillFields)
    console.log(surveyLink)
  })
}

function test_ParticipantsData()
{
  console.log(JSON.stringify(
    getParticipantsData(getCohortSettings())
    , null, 2
  ))

  console.log(JSON.stringify(
    getParticipantsData(getCohortSettings(), { getBlanks: true })
    , null, 2
  ))
}
