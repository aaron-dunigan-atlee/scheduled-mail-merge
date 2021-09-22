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
