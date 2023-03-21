function include(filename)
{
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
};


function getAccessToken()
{
  var token = ScriptApp.getOAuthToken();

  return token // DriveApp.getFiles()
}

function test_businessDays()
{
  var date = new Date('07/29/2020')
  console.log(date)
  console.log(addBusinessDays(date, -1))
}


/**
 * Calculate number of business days (may be negative for days before) from a date, 
 * including common US holidays.
 * Adapted from a few answers at https://stackoverflow.com/questions/21297323/calculate-an-expected-delivery-date-accounting-for-holidays-in-business-days-u
 * @param {Date} date 
 * @param {number} businessDays 
 */
function addBusinessDays(date, businessDays)
{
  if (!date || !(date instanceof Date))
  {
    console.log("Can't add business days to " + date + " because it's not a date.")
    return date;
  }
  console.log("Adding " + businessDays + ' business days to ' + date)
  var tmp = new Date(date);

  // Timezones are scary, let's work with whole-days only
  if (businessDays !== parseInt(businessDays, 10))
  {
    throw new TypeError('businessDaysFromDate can only adjust by whole days');
  }

  // short-circuit no work; make direction assignment simpler
  if (businessDays === 0)
  {
    return date;
  }

  var direction = businessDays > 0 ? 1 : -1;
  businessDays = Math.abs(businessDays)

  while (businessDays > 0)
  {
    tmp.setDate(tmp.getDate() + direction);
    if (isBusinessDay(tmp))
    {
      console.log(tmp + ' is a business day')
      --businessDays;
    } else
    {
      console.log(tmp + ' is not a business day')
    }
  }
  console.log('Final date: ' + tmp)
  return tmp;

  // Private functions
  // -----------------

  /**
   * Taken from https://stackoverflow.com/a/39455759
   * As noted in the comments there, it may not work perfectly.
   * @param {Date} date 
   */
  function isBusinessDay(date)
  {
    var dayOfWeek = date.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6)
    {
      // Weekend
      return false;
    }

    holidays = [
      '12/31+5', // New Year's Day on a saturday celebrated on previous friday
      '1/1',     // New Year's Day
      '1/2+1',   // New Year's Day on a sunday celebrated on next monday
      '1-3/1',   // Birthday of Martin Luther King, third Monday in January
      '2-3/1',   // Washington's Birthday, third Monday in February
      '5~1/1',   // Memorial Day, last Monday in May
      '7/3+5',   // Independence Day
      '7/4',     // Independence Day
      '7/5+1',   // Independence Day
      '9-1/1',   // Labor Day, first Monday in September
      '10-2/1',  // Columbus Day, second Monday in October
      '11/10+5', // Veterans Day
      '11/11',   // Veterans Day
      '11/12+1', // Veterans Day
      '11-4/4',  // Thanksgiving Day, fourth Thursday in November
      '12/24+5', // Christmas Day
      '12/25',   // Christmas Day
      '12/26+1',  // Christmas Day
    ];

    var dayOfMonth = date.getDate(),
      month = date.getMonth() + 1,
      monthDay = month + '/' + dayOfMonth;

    if (holidays.indexOf(monthDay) > -1)
    {
      return false;
    }

    var monthDayDay = monthDay + '+' + dayOfWeek;
    if (holidays.indexOf(monthDayDay) > -1)
    {
      return false;
    }

    var weekOfMonth = Math.floor((dayOfMonth - 1) / 7) + 1,
      monthWeekDay = month + '-' + weekOfMonth + '/' + dayOfWeek;
    if (holidays.indexOf(monthWeekDay) > -1)
    {
      return false;
    }

    var lastDayOfMonth = new Date(date);
    lastDayOfMonth.setMonth(lastDayOfMonth.getMonth() + 1);
    lastDayOfMonth.setDate(0);
    var negWeekOfMonth = Math.floor((lastDayOfMonth.getDate() - dayOfMonth - 1) / 7) + 1,
      monthNegWeekDay = month + '~' + negWeekOfMonth + '/' + dayOfWeek;
    if (holidays.indexOf(monthNegWeekDay) > -1)
    {
      return false;
    }

    return true;
  } // addBusinessDays.isBusinessDay()

} // addBusinessDays()

/**
 * Calculate number business days (may be negative for days before) from a date, 
 * @param {Date} date 
 * @param {number} daysDiff 
 */
function addDays(date, daysDiff)
{
  if (!date || !(date instanceof Date))
  {
    console.log("Can't add days to " + date + " because it's not a date.")
    return date;
  }
  console.log("Adding " + daysDiff + ' business days to ' + date)
  var tmp = new Date(date);

  // Timezones are scary, let's work with whole-days only
  if (daysDiff !== parseInt(daysDiff, 10))
  {
    throw new TypeError('addDays can only adjust by whole days');
  }

  // short-circuit no work; make direction assignment simpler
  if (daysDiff === 0)
  {
    return date;
  }

  var direction = daysDiff > 0 ? 1 : -1;
  daysDiff = Math.abs(daysDiff)

  while (daysDiff > 0)
  {
    tmp.setDate(tmp.getDate() + direction);
    --daysDiff;
  }
  console.log('Final date: ' + tmp)
  return tmp;
}
/**
 * Convert a date string to a Date object with correct conversion to local timezone 
 * @param {string} dateString A valid date string, expressed in the local timezone 
 */
function stringToLocalDate(dateString)
{
  if (!dateString) return null;
  try
  {
    var date = new Date(dateString)
    date.setMinutes(date.getMinutes() + date.getTimezoneOffset())
    return date
  } catch (err)
  {
    return null;
  }
}

/**
 * @returns {Object} Keys are column A (NOT normalized); values are Column B
 */
function getSettings()
{
  // Get the data we'll need to fill the global values in the templates.
  var ss = SpreadsheetApp.getActive()
  var settingsArray = ss.getSheetByName('⚙ Settings').getDataRange().getValues()
  var cohortSettings = {}
  settingsArray.forEach(function (row) { if (row[0]) cohortSettings[row[0]] = row[1] })
  return cohortSettings
}

function shareSilentyFailSilently(fileId, userEmail, role)
{
  role = role || 'reader'
  try
  {
    Drive.Permissions.insert(
      {
        'role': role,
        'type': 'user',
        'value': userEmail
      },
      fileId,
      {
        'sendNotificationEmails': 'false'
      });
  } catch (err)
  {
    slackCacheWarn("Couldn't share file " + fileId + " with " + userEmail + ": " + err.message)
  }
}