function OBJ(SHEET_ID,SHEET_NAME,RANGE) {
  
    var arrayOfObjects = CreateOBJ(SHEET_ID, SHEET_NAME, RANGE);
    // Log the array of objects
    //Logger.log(JSON.stringify(arrayOfObjects));
    //writeObjToScriptProperties('OBJ', arrayOfObjects,SCRIPT_PROPERTIES_SERVICE);
    //Logger.log(arrayOfObjects);
    return arrayOfObjects;
  
  }
  
  function CreateOBJ(SHEET_ID, SHEET_NAME, RANGE) {
    var sheet = SpreadsheetApp.openById(SHEET_ID);
    var dataSheet = sheet.getSheetByName(SHEET_NAME);
    var range = dataSheet.getRange(RANGE);
    var values = range.getValues();
    var arrayOfObjects = [];
    var headers = values[0];
    var index = 1;
    var columnValues = getColumnValues(index);
  
    var lastRow = sheet.getLastRow();
    writeValueToScriptProperties('ROW', lastRow, SCRIPT_PROPERTIES_SERVICE);
  
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var dataObject = {};
  
      for (var j = 0; j < row.length; j++) {
        var columnName = headers[j];
        var columnValue = row[j];
        dataObject[columnName] = columnValue;
  
      }
  
      var email = columnValues[i];
  
      // Adjusting the structure of arrayOfObjects
      arrayOfObjects.push( 
        { Email : email, Info: dataObject});
    }
  
    return arrayOfObjects ;
  }
  
  function getColumnValues(index) {
    // Access the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
    // Access the active sheet
    var sheet = spreadsheet.getActiveSheet();
  
    // Define the column you want to retrieve values from (let's say column A)
    var columnNumber = index; // Column A is the 1st column
  
    // Get the data range of the entire column
    var columnRange = sheet.getRange(1, columnNumber, sheet.getLastRow(), 1);
  
    // Get the values in the column
    var columnValues = columnRange.getValues();
  
    // Log the values (this will be visible in the Apps Script Execution logs)
    //Logger.log(columnValues);
  
    // Flatten the 2D array to a 1D array
    return columnValues.map(function (value) {
      return value[0];
    });
  }
  
  