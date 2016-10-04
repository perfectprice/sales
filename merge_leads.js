/* 
A simple script to merge a list of leads from sheet 1 (2nd tab) into Sheet 0 (1st tab)
Unique key is email
New leads that do not exist in Sheet 0 are appended at the end
Written by Alexander Shartsis github: @thegeneralist
(With lots of help from Google's tutorials, because I don't really know javascript)
*/

// Lists must have same columns in the same order, though blanks are fine.

// Constants

var SOURCE_UNIQUE_KEY = 1; // index key for lists to merge together. This is Hubspot CRM's key and therefore a good choice as a constant.
var TARGET_UNIQUE_KEY = 1; // index key for lists to merge together. This is Hubspot CRM's key and therefore a good choice as a constant.
var KEY = "email";

function mergeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var target = ss.getSheets()[0];
  var targetDataRange = target.getRange(2, 1, target.getMaxRows() -1, target.getMaxColumns());
  targetDataRange.sort(TARGET_UNIQUE_KEY);
  var targetObjects = getRowsData(target, targetDataRange, 1);
  
  var source = ss.getSheets()[1];
  var sourceDataRange = source.getRange(2, 1, source.getMaxRows() -1, source.getMaxColumns());
  sourceDataRange.sort(SOURCE_UNIQUE_KEY);
  var sourceObjects = getRowsData(source, sourceDataRange, 1);

  var newList = {};
  
  newList = mergeLists(sourceObjects.objects, targetObjects.objects, KEY);
  var arrayToWrite = makeArray(newList, targetObjects.headers);
  
  arrayDimensions = findDimensions(arrayToWrite);
  
  ss.insertSheet();
  sheet = ss.getActiveSheet();
  range = sheet.getRange(1, 1, arrayDimensions[0], arrayDimensions[1]);                        
  range.setValues(arrayToWrite); 
}


function makeArray(objects, headers) {
  var values = [];
  values[0] = headers;
  for (var i = 0; i < objects.length; ++i) {
    var rowValues = [];
    for (var j = 0; j < headers.length; ++j) {
      rowValues.push(objects[i][headers[j]] || "");
    }
    values.push(rowValues);
  }
  return values;
}

// Iterate through the source and target lists
// If there is a match, merge new source data with target data (source overwrites target if nonblank!)
// If no match, add a new object
// Returns merged objects array
// Arguments:
//     - sourceObjects: 3d javascript array of objects that is new to be merged IN, will overwrite if non-blank
//     - targetObjects: 3d javascript array of objects to be merged into, presmuable older and therefore can be overwritten in conflict


// TO DO: Update w/ headers passing through and being returned back

function mergeLists(sourceObjects, targetObjects, key) {
  var mergedObjects = [];
  
  do {
    if (sourceObjects[0][key] == targetObjects[0][key]) {

      // Join all properties into targetObjects, overwriting targetObject if a collision
      for (var propertyName in sourceObjects[0]) {
        targetObjects[0][propertyName] = sourceObjects[0][propertyName];
      }
      
      // push the targetObject that is the joined version
      mergedObjects.push(targetObjects[0]);
      sourceObjects.shift();
      targetObjects.shift();
    }
    else if (sourceObjects[0][key] < targetObjects[0][key]) {   // source should be first
      mergedObjects.push(sourceObjects[0]);
      sourceObjects.shift();
    } 
    else 
    {              // if source shouldn't be first, then target should be first
      mergedObjects.push(targetObjects[0]);
      targetObjects.shift();
    }
  } while (targetObjects.length != 0 && sourceObjects.length != 0);
  
  
  if (sourceObjects.length > 0 && targetObjects.length == 0) {
    do {
      mergedObjects.push(sourceObjects[0]);
      sourceObjects.shift();
    } while (sourceObjects.length > 0);
  } 
  else if (targetObjects.length > 0 && sourceObjects.length == 0) {
     do {
        mergedObjects.push(targetObjects[0]);
        targetObjects.shift();
      } while (targetObjects.length > 0);
    } 

  return mergedObjects;
}

function findDimensions(a){
    var mainLen = 0;
    var subLen = 0;

    mainLen = a.length;

    for(var i=0; i < mainLen; i++){
        var len = a[i].length;
        subLen = (len > subLen ? len : subLen);
    }

    return [mainLen, subLen];
};


// returns just the headers from a sheet
// if columnHeadersRowIndex is not given, it is assumed to be row 1
// normalizes them using normalizeHeaders()

function getHeaders(sheet, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || 0;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return normalizeHeaders(headers);
}


//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return { objects: getObjects(range.getValues(), normalizeHeaders(headers)), headers : normalizeHeaders(headers) };
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}
