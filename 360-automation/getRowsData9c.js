/**  
 * getRowsData9c.gs requires getRowsData.gs
 * version 9c: added preserveArrayFormulas to parameters
 *             setRowsData2 now returns gracefully if object array is empty 
 */

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below headersRowIndex (if defined).
//   - parameters 
//     headersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
//     getDisplayValues: gets the display values as strings
//   - headersCase: the case of the returned property values, default is lowerCamelCase (camelCase,snake_case,lowercase)
//
// Returns an Array of objects.
//
function getRowsData2(sheet, range, parameters) {
  parameters = parameters || {}
  if (sheet.getLastRow() < 2) return [];
  var headersIndex = parameters.headersRowIndex || (range ? range.getRowIndex() - 1 : 1);
  
  var dataRange, base;
  if (!range && (parameters.startHeader || parameters.endHeader)){
    range = getBodyRange(sheet, headersIndex, parameters.startHeader,parameters.endHeader);  
  }
  if (range){
    dataRange = range;
  } else {
    var numRows = sheet.getLastRow() - headersIndex;
    if (numRows <= 0) return [];
    dataRange = sheet.getRange(headersIndex+1, 1, numRows, sheet.getLastColumn());
  }
  
  var numColumns = dataRange.getLastColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  var values = (parameters.getDisplayValues || parameters.displayValues) ? dataRange.getDisplayValues() : dataRange.getValues();
  //backgrounds, notes
  var keys;
  if (!parameters.headersCase || parameters.headersCase === 'camelCase' || parameters.headersCase === 'camel') keys = normalizeHeaders(headers);
  if (parameters.headersCase === 'snake_case' || parameters.headersCase === 'snake') keys = snakeCaseHeaders(headers);
  if (parameters.headersCase === 'lowercase' || parameters.headersCase === 'lower') keys = lowerCaseHeaders(headers);
  if (parameters.getShortcut) base = sheet.getParent().getUrl()+'#gid='+sheet.getSheetId()+'&range=A';
  
  return getObjects_(values, keys, parameters.getBlanks, parameters.getMetadata, dataRange.getRowIndex(), base);
}


function snakeCaseHeaders(headers){
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    keys.push(headers[i].replace(/\W/g,'_').toLowerCase());
  }
  return keys;
}

function lowerCaseHeaders(headers){
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    keys.push(headers[i].replace(/\W/g,'').toLowerCase());
  }
  return keys;
}


// Parameters:
//   - writeMethod:
//       overwrite (default): writes the data regardless of what is already present in the range
//       append: adds the new rows afer the last row with data on the sheet
//       clear: writes the new rows, then clears all rows beneath the destination range
//       delete: writes the new rows, then deletes all rows beneath the destination range
//   - headersRowIndex (integer): index where the column headers are defined. This defaults to the row 1.
//   - startHeader (string): will look for an exact match to be used leftmost bound of the range where data will be written, throws error if not found
//   - endHeader (string): will look for an exact match to be used rightmost bound of the range where data will be written, throws error if not found
//   - omitZeros: if true, cells with the value of zero will be omitted, writing blanks instead
//   - firstRowIndex (integer): index of the first row where data should be written. This defaults to the row immediately below the headers.
//   - headersCase: how to transform the case of the headers (defaults to camelCase), default is lowerCamelCase (camelCase,snake_case,lowercase,none)
//   - preserveArrayFormulas: If true, don't overwrite cell if its row 1 has an array formula.

function setRowsData2(sheet, objects, parameters) {
  if (objects.length === 0) {
    console.log("setRowsData2: No data to write.")
    return
  }
  parameters = parameters || {};
  var writeMethod = parameters.writeMethod || 'overwrite';
  var headersRowIndex = parameters.headersRowIndex || 1;
  var headersRange = getHeadersRange(sheet,headersRowIndex,parameters.startHeader,parameters.endHeader);
  //console.log('Headers range is '+headersRange.getA1Notation());
  
  var firstRowIndex;
  if (writeMethod === 'append') firstRowIndex = sheet.getLastRow()+1;
  if (writeMethod === 'overwrite' || writeMethod === 'clear' || writeMethod === 'delete') firstRowIndex = headersRange.getRowIndex() + 1;
  if (parameters.firstRowIndex) firstRowIndex = parameters.firstRowIndex;
  
  var headers = headersRange.getValues().shift();
  var keys;
  if (!parameters.headersCase || parameters.headersCase === 'camelCase' || parameters.headersCase === 'camel') keys = normalizeHeaders(headers);
  if (parameters.headersCase === 'snake_case' || parameters.headersCase === 'snake') keys = snakeCaseHeaders(headers);
  if (parameters.headersCase === 'lowercase' || parameters.headersCase === 'lower') keys = lowerCaseHeaders(headers);
  if (parameters.headersCase === 'none') keys = headers;
  //backgrounds, notes
  if (!objects instanceof Array && objects instanceof Object) objects = [objects]; //in case only one object is passed instead of an array with one element as intended
  var formulaKeys = {};
  if (parameters.preserveArrayFormulas) {
    var headerFormulas = sheet.getRange(1, headersRange.getColumn(), 1, headersRange.getLastColumn()).getFormulas().shift();
    for (j = 0; j < keys.length; ++j) {
      if (headerFormulas[j]) formulaKeys[keys[j]]=true;
    }
  }

  
  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < keys.length; ++j) {
      var header = keys[j];
      if (header.length > 0){
        if (parameters.preserveArrayFormulas && formulaKeys[header]) {
          values.push(null);
        } else if (parameters.omitZeros || parameters.omitZeroes){
          values.push(objects[i][header] ? objects[i][header] : "");
        } else {
          values.push(typeof objects[i][header] !== 'undefined' ? objects[i][header] : ""); //what about null
        }
      } else { //else column header is blank
        values.push("")
      }
    }
    data.push(values);
  }
  
  var destinationRange = sheet.getRange(firstRowIndex, headersRange.getColumnIndex(), objects.length, headers.length);
  if (writeMethod === 'clear' && sheet.getLastRow() - destinationRange.getLastRow() > 0){
    var clearRange = sheet.getRange(destinationRange.getLastRow()+1,1,sheet.getLastRow() - destinationRange.getLastRow(),sheet.getLastColumn())
    console.log('Cleared range: '+clearRange.getA1Notation());
    clearRange.clear()
  }
  if (writeMethod === 'delete' && sheet.getMaxRows() - destinationRange.getLastRow() > 0){
    var firstRowToDelete = destinationRange.getLastRow()+1;
    var numRowsToDelete = sheet.getMaxRows() - destinationRange.getLastRow();
    console.log('Deleted '+firstRowToDelete+' rows starting on row '+numRowsToDelete+'.');
    sheet.deleteRows(firstRowToDelete,numRowsToDelete);
  }
  destinationRange.setValues(data);
  return destinationRange
}



//Helper function that gets the headers range, optionally matching header values to determine start and end 
function getHeadersRange(sheet,headersRowIndex,startHeader,endHeader){
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(headersRowIndex,1,1,lastCol).getValues().shift();
  var columnBounds = getStartEndColumns(headers,startHeader,endHeader,lastCol)
  lastCol = columnBounds.endCol - columnBounds.startCol+1;
  var headersRange = sheet.getRange(headersRowIndex, columnBounds.startCol, 1, lastCol);
  return headersRange;
}

//Helper function that gets the body range, optionally matching header values to determine start and end 
function getBodyRange(sheet,headersRowIndex,startHeader,endHeader){
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(headersRowIndex,1,1,lastCol).getValues().shift();
  var columnBounds = getStartEndColumns(headers,startHeader,endHeader,lastCol)
  lastCol = columnBounds.endCol - columnBounds.startCol+1;
  var numRows = sheet.getLastRow() - headersRowIndex;
  var bodyRange = sheet.getRange(headersRowIndex+1, columnBounds.startCol, numRows, lastCol);
  return bodyRange;
}

function getStartEndColumns(headers,startHeader,endHeader,lastCol){
  if (!endHeader) var endCol = lastCol
  if (endHeader) {
    var endCol = headers.indexOf(endHeader)+1;
    if (!endCol){
      throw 'endHeader "'+endHeader+'" column not found';
    }
  }
  if (!startHeader) var startCol = 1;
  if (startHeader){
    var startCol = headers.indexOf(startHeader)+1;
    if (!startCol){
      throw 'startHeader "'+startHeader+'" column not found';
    }
  }
  return {startCol:startCol,endCol:endCol}
}