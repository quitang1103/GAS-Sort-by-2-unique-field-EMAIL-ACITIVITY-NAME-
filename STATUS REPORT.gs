function RUN2(){
    var name = "STATUS_REPORT";
    ProgressStatus(name);
  }
  
  
  function ProgressStatus(name) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var statusSheet = spreadsheet.getSheetByName(name);
    var startTime2 = new Date().getTime();
  
    // Check if 'STATUS_REPORT' exists, and create it if not
    if (!statusSheet) {
      spreadsheet.insertSheet(name);
      var statusSheet = spreadsheet.getSheetByName(name); // Retrieve the sheet after creation
    } 
    var exportSheet = spreadsheet.getSheetByName('EXPORT');
    //var dataRange = exportSheet.getDataRange(); // Get the entire data range in the 'EXPORT' sheet
  
    var numRows = exportSheet.getLastRow();
    var numColumns = exportSheet.getLastColumn() ;
    var startRow = 1;
    var startColumn = 1;
    var range = exportSheet.getRange(startRow + 1, startColumn +1 , numRows, numColumns);
    var range2 = exportSheet.getRange(startRow + 1, startColumn +1 , numRows, numColumns);
    
    var values = range.getValues(); // Get all values in the data range
     var values2 = range2.getValues();
    Logger.log(values[0]);
  
    //Logger.log('Original Array of Arrays: ' + JSON.stringify(values));
    var myArray = switchValuesInSubarrays(values, 1, 2);
    var myArray2 = switchValuesInSubarrays(values2, 1, 2);
    //Logger.log('Array of Arrays after switching positions: ' + JSON.stringify(myArray));
    Logger.log("+++++1+" + myArray[0]);
    Logger.log("+++++2+" + myArray2[0]);
  
    //Logger.log('Original Array of Arrays: ' + JSON.stringify(myArray));
    myArray = removeColumnRangeInSubarrays(myArray,2,8);
    myArray2 = clearSpecificColumnRangeInSubarrays(myArray2, 2, 7);
    Logger.log('Array of Arrays after clearing specific column range: ' + JSON.stringify(myArray));
    //Logger.log('Array of Arrays after clearing specific column2 range: ' + JSON.stringify(myArray2));
    var resultMatrix = removeDuplicateSubarrays(1,myArray);
    var resultMatrix2 = removeDuplicateSubarrays(1,myArray2);
    Logger.log("+++++++++++++: " +resultMatrix[0]);
    statusSheet.getRange(2, 1, resultMatrix.length, resultMatrix[0].length).setValues(resultMatrix);
  
    // Corrected headers array with string values
    var headers = [['Email','Họ & Tên','TEST', 'Sản phẩm', 'KHBD', 'Slide duyệt giảng', 'Tự luyện giảng', 'Duyệt giảng']];
    
    // Set headers in 'statusSheet'
    statusSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    
    var cell = statusSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    cell.setBackground('#5dc2a7');
    cell.setFontWeight('bold');
    cell.setFontSize(14)
    cell.setHorizontalAlignment('center');
    cell.setVerticalAlignment('middle');
    cell.setFontColor('#FFFFFF');
    cell.setBorder(true, true, true, true, true, true);
  
    for (var i = 1; i <= headers[0].length; i++) {
      statusSheet.autoResizeColumn(i);
    }
  
    //var Unique_Email = getUniqueEmail(2,"EXPORT");
    //Logger.log("11111111111111"+ Unique_Email);
    var Unique_Activities = getUniqueActivities(3,"EXPORT");
    //Logger.log("33333333333333"+JSON.stringify(Unique_Activities));
    //var Unique_NAME = getUniqueActivities(4,"EXPORT");
    //Logger.log("44444444444"+JSON.stringify(Unique_NAME));
    removeDuplicateRows(1,'STATUS_REPORT');
    var result = [];
  
    var flag2 = readValueFromScriptProperties('FLAG2', SCRIPT_PROPERTIES_SERVICE);
    if (flag2 === null || flag2 === undefined) {
      flag2 = 0;
    }
  
    for (var i = flag2; i < resultMatrix2.length; i++) {
      var row = resultMatrix2[i];
      var flag = flag2
      Logger.log('MAX :' + resultMatrix2.length );
  
      if (shouldExit(startTime2)) {
            Logger.log('Elapsed time exceeded. Scheduling trigger to resume.');
            writeValueToScriptProperties('FLAG2', flag, SCRIPT_PROPERTIES_SERVICE);
            var startTime = new Date().getTime();
            return;
          } else {
            
  
      for (var c = 2; c < row.length; c++) {
        Logger.log(row.length)
    // Perform your specific activity here
    // For example, set a new value for the cell
  
    switch (c) {
      case 2:
        // Code to execute when c is 2
        row[c] = "1New Value";
        var startRow = 2; // Adjust the starting row as needed
        var startColumn = 1; // Adjust the starting column as needed
        var cellRow = startRow + i ;
        var cellColumn = startColumn + c; // Adjust for 1-based indexing
        var statusSheet = spreadsheet.getSheetByName('STATUS_REPORT');
        var RANGE = statusSheet.getName() + '!' + String.fromCharCode(64 + cellColumn) + cellRow;
        Logger.log("++++++++++++"+RANGE);
        Logger.log("++++++++++++"+statusSheet);
        SCROLL_DOWN(RANGE, statusSheet.getName());
        c++;
  
      case 3:
        // Code to execute when c is 3
        row[c] = "2New Value";
        var startRow = 2; // Adjust the starting row as needed
        var startColumn = 1; // Adjust the starting column as needed
        var cellRow = startRow + i ;
        var cellColumn = startColumn + c; // Adjust for 1-based indexing
        var statusSheet = spreadsheet.getSheetByName('STATUS_REPORT');
        var RANGE = statusSheet.getName() + '!' + String.fromCharCode(64 + cellColumn) + cellRow;
        Logger.log("++++++++++++"+RANGE);
        Logger.log("++++++++++ : c "+ c);
        SCROLL_DOWN(RANGE, statusSheet.getName());
        c++;
      case 4:
        // Code to execute when c is 4
        row[c] = "3New Value";
        var startRow = 2; // Adjust the starting row as needed
        var startColumn = 1; // Adjust the starting column as needed
        var cellRow = startRow + i ;
        var cellColumn = startColumn + c; // Adjust for 1-based indexing
        var statusSheet = spreadsheet.getSheetByName('STATUS_REPORT');
        var RANGE = statusSheet.getName() + '!' + String.fromCharCode(64 + cellColumn) + cellRow;
        Logger.log("++++++++++++"+RANGE);
        Logger.log("++++++++++ : c "+ c);
         SCROLL_DOWN(RANGE, statusSheet.getName());
        c++;
      case 5:
        // Code to execute when c is 5
        row[c] = "4New Value";
        var startRow = 2; // Adjust the starting row as needed
        var startColumn = 1; // Adjust the starting column as needed
        var cellRow = startRow + i ;
        var cellColumn = startColumn + c; // Adjust for 1-based indexing
        var statusSheet = spreadsheet.getSheetByName('STATUS_REPORT');
        var RANGE = statusSheet.getName() + '!' + String.fromCharCode(64 + cellColumn) + cellRow;
        Logger.log("++++++++++++"+RANGE);
        Logger.log("++++++++++ : c "+ c);
        SCROLL_DOWN(RANGE, statusSheet.getName());
        c++;
      case 6:
        // Code to execute when c is 6
        row[c] = "5New Value";
        var startRow = 2; // Adjust the starting row as needed
        var startColumn = 1; // Adjust the starting column as needed
        var cellRow = startRow + i ;
        var cellColumn = startColumn + c; // Adjust for 1-based indexing
        var statusSheet = spreadsheet.getSheetByName('STATUS_REPORT');
        var RANGE = statusSheet.getName() + '!' + String.fromCharCode(64 + cellColumn) + cellRow;
        Logger.log("++++++++++++"+RANGE);
        Logger.log("++++++++++ : c "+ c);
        SCROLL_DOWN(RANGE, statusSheet.getName());
        c++;
      case 7:
        // Code to execute when c is 7
        row[c] = "6New Value";
        var startRow = 2; // Adjust the starting row as needed
        var startColumn = 1; // Adjust the starting column as needed
        var cellRow = startRow + i ;
        var cellColumn = startColumn + c; // Adjust for 1-based indexing
        var statusSheet = spreadsheet.getSheetByName('STATUS_REPORT');
        var RANGE = statusSheet.getName() + '!' + String.fromCharCode(64 + cellColumn) + cellRow;
        Logger.log("++++++++++++"+RANGE);
        Logger.log("++++++++++ : c "+ c);
        SCROLL_DOWN(RANGE, statusSheet.getName());
        c++;
      default:
        // Code to execute when c is not 2, 3, 4, or 5 ,6 ,7 
        row[c] = "Default New Value";
        var startRow = 2; // Adjust the starting row as needed
        var startColumn = 1; // Adjust the starting column as needed
        var cellRow = startRow + i ;
        var cellColumn = startColumn + c; // Adjust for 1-based indexing
        var statusSheet = spreadsheet.getSheetByName('STATUS_REPORT');
        var RANGE = statusSheet.getName() + '!' + String.fromCharCode(64 + cellColumn) + cellRow;
        Logger.log("++++++++++++"+RANGE);
        Logger.log("++++++++++ : c "+ c);
        break;}
        flag2++;
        writeValueToScriptProperties('FLAG2', flag2, SCRIPT_PROPERTIES_SERVICE);
  
    }
  
  
      }
      //result.push(row);
    }
  
    // Update the entire range in the 'EXPORT' sheet with the modified resultMatrix
    
    //result = removeColumn(result,7);
    //range2 = statusSheet.getRange(2,1,result.length,result[0].length);
    //range2.setValues(result);
    return
    }
  
  
  
  
  
  function SCROLL_DOWN(RANGE,name,condition,data1,data2) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(name);
    var condition = condition;
    var data1 = data1;
    var data2 = data2;
  
    // Dữ liệu cho trình đơn thả xuống và màu sắc tương ứng
    var dataWithColors = [
      { value: 'SUBMITTED', color: '#87CEFA' },  // 
      { value: 'PASS', color: '#00FF00' },       // Green
      { value: 'FAIL', color: '#FF0000' }        // Red
    ];
  
    // Cell để chèn trình đơn thả xuống
    var cell = sheet.getRange(RANGE); // Điều chỉnh tham chiếu ô theo ý muốn
  
    // Xóa dữ liệu hiện tại trong ô
    cell.clearContent();
  
  
  
  
  
  
    // Thiết lập kiểm tra dữ liệu để tạo trình đơn thả xuống
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(dataWithColors.map(item => item.value), true).build();
    cell.setDataValidation(rule);
  
    // Thiết lập màu nền và viền cho từng lựa chọn
    dataWithColors.forEach(function (item, index) {
      var optionRange = cell.offset(index, 0, 1, 1);
      optionRange.setValue(item.value); // Set the text value
      optionRange.setBackground(item.color);
      optionRange.setBorder(true, true, true, true, true, true, item.color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    });
  }
  
  
  
  
  function switchArrayPositions(array, position1, position2) {
    // Check if the positions are within the array bounds
    if (position1 >= 0 && position1 < array.length && position2 >= 0 && position2 < array.length) {
      // Swap the values at the specified positions
      var temp = array[position1];
      array[position1] = array[position2];
      array[position2] = temp;
    } else {
      console.error('Invalid positions provided');
    }
    return array;
  }
  
  function switchValuesInSubarrays(arrayOfArrays, position1, position2) {
    var modifiedArray = [];
    for (var i = 0; i < arrayOfArrays.length; i++) {
      var subarray = arrayOfArrays[i];
  
      // Ensure the subarray has at least two elements
      if (subarray.length >= Math.max(position1, position2) + 1) {
        // Switch the positions of the values
        var temp = subarray[position1];
        subarray[position1] = subarray[position2];
        subarray[position2] = temp;
      }
      modifiedArray.push(subarray);
    }
    return modifiedArray;
  }
  
  
  function clearSpecificColumnRangeInSubarrays(arrayOfArrays, startColumn, endColumn) {
    // Create a new array to store the modified subarrays
    var modifiedArray = [];
  
    for (var i = 0; i < arrayOfArrays.length; i++) {
      var subarray = arrayOfArrays[i].slice(); // Create a copy of the subarray
  
      // Ensure the subarray has the specified range of columns
      if (subarray.length >= endColumn + 1) {
        // Clear the specific range of columns in each subarray
        for (var col = startColumn; col <= endColumn; col++) {
          subarray[col] = '';
          // Alternatively, you can set the specific column to another value of your choice
        }
      }
  
      // Add the modified subarray to the result array
      modifiedArray.push(subarray);
    }
  
    // Return the modified array
    return modifiedArray;
  }
  
  function getUniqueEmail(index,name) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(name); // Replace 'Sheet1' with your sheet name
    var activityColumn = index;
    var lastRow = sheet.getLastRow();
    var activityRange = sheet.getRange(2, activityColumn, lastRow - 1, 1); // Assuming data starts from row 2
    var activityValues = activityRange.getValues();
  
    // Create an object to track unique activity names
    var activityMap = {};
    var uniqueActivities = [];
  
    // Process each row
    for (var i = 0; i < activityValues.length; i++) {
      var activity = activityValues[i][0];
  
      // Check if the activity is already in the map
      if (activityMap[activity]) {
        // Activity is a duplicate, clear the value
        sheet.getRange(i + 2, activityColumn).clearContent();
      } else {
        // Activity is unique, update the map and add to the uniqueActivities array
        activityMap[activity] = true;
        uniqueActivities.push(activity);
      }
    }
  
    // Log the array of unique activities
    Logger.log('Unique Email: DONE');
    return uniqueActivities;
  }
  
  function getUniqueActivities(index,name) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(name); // Replace 'Sheet1' with your sheet name
  
    // Assuming activity names are in column B
    var activityColumn = index;
  
    var lastRow = sheet.getLastRow();
    var activityRange = sheet.getRange(2, activityColumn, lastRow - 1, 1); // Assuming data starts from row 2
  
    var activityValues = activityRange.getValues();
  
    // Create an object to track unique activity names
    var activityMap = {};
    var uniqueActivities = [];
  
    // Process each row
    for (var i = 0; i < activityValues.length; i++) {
      var activity = activityValues[i][0];
  
      // Check if the activity is already in the map
      if (!activityMap[activity]) {
        // Activity is unique, update the map and add to the uniqueActivities array
        activityMap[activity] = true;
        uniqueActivities.push(activity);
      }
    }
  
    // Log the array of unique activities
    Logger.log('Unique Activities: DONE');
  
    return uniqueActivities;
  }
  
  
  function removeDuplicateRows(index,name) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName(name); // Replace 'Sheet1' with your sheet name
  
    // Assuming the column with values to check for duplicates is column B
    var columnToCheck = index;
  
    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues(); // Assuming data starts from row 2
  
    var uniqueValues = {};
    var duplicateRows = [];
  
    // Process each row
    for (var i = 0; i < values.length; i++) {
      var value = values[i][columnToCheck - 1]; // Adjust for 0-based index
  
      // Check if the value is already in the map
      if (!uniqueValues[value]) {
        // Value is unique, update the map
        uniqueValues[value] = true;
      } else {
        // Value is a duplicate, add the row index to the duplicateRows array
        duplicateRows.push(i + 2); // Adding 2 to get the correct sheet row index
      }
    }
  
    // Remove duplicate rows
    for (var j = duplicateRows.length - 1; j >= 0; j--) {
      sheet.deleteRow(duplicateRows[j]);
    }
  }
  
  function removeDuplicateSubarrays(index, matrix) {
    // Create an object to track unique values
    var uniqueValues = {};
    var uniqueMatrix = [];
  
    // Process each subarray
    for (var i = 0; i < matrix.length; i++) {
      var value = matrix[i][index];
  
      // Check if the value is already in the map
      if (!uniqueValues[value]) {
        // Value is unique, update the map and add the subarray to the uniqueMatrix
        uniqueValues[value] = true;
        uniqueMatrix.push(matrix[i]);
      }
      // If the value is a duplicate, skip adding the subarray
    }
  
    return uniqueMatrix;
  }
  
  function removeColumn(array, columnIndexToRemove) {
    return array.map(row => row.filter((_, index) => index !== columnIndexToRemove));
  }
  
  
  function removeColumnRangeInSubarrays(arrayOfArrays, startColumnIndex, endColumnIndex) {
    for (var i = 0; i < arrayOfArrays.length; i++) {
      var subarray = arrayOfArrays[i];
  
      // Check if the column indices are valid
      if (
        startColumnIndex >= 0 &&
        startColumnIndex < subarray.length &&
        endColumnIndex >= startColumnIndex &&
        endColumnIndex < subarray.length
      ) {
        // Remove the range of elements from startColumnIndex to endColumnIndex
        subarray.splice(startColumnIndex, endColumnIndex - startColumnIndex + 1);
      }
    }
  
    return arrayOfArrays;
  }
  
  function shouldExit(startTime2) {
    const elapsedMillis = new Date().getTime() - startTime2;
    const elapsedMinutes = elapsedMillis / (1000 * 60); // Convert milliseconds to minutes
  
    // Check if elapsed time exceeds the threshold (e.g., 4 minutes)
    return elapsedMinutes > 4.3;
  }
  
  
  function clearFlag2(){
  clearScriptProperty('FLAG2');
  }
  
  function initializeFlag2(){
    value = 0;
    writeValueToScriptProperties('FLAG2',value,SCRIPT_PROPERTIES_SERVICE);
  }
  
  function removeSTATUS_REPORT() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('STATUS_REPORT');
    if (sheet) {
      // If the sheet exists, delete it
      spreadsheet.deleteSheet(sheet);
      Logger.log('Sheet "' + sheetName + '" deleted successfully.');
    } else {
      Logger.log('Sheet "' + sheetName + '" does not exist.');
    }
  }
  
  
  function compare(data1,data2,index){
  
    
  
  }
  
  
  
  