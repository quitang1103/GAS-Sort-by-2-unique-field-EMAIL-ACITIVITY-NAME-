function exportMergeDataToSheet(exportSheet, data) {
    var newSheet = exportSheet;
    var flag = readValueFromScriptProperties('FLAG', SCRIPT_PROPERTIES_SERVICE);
    var limit = readValueFromScriptProperties('ROW', SCRIPT_PROPERTIES_SERVICE);
  
    // Record the start time
    var startTime = new Date().getTime();
  
    // Check if data is not undefined or null
    if (!data) {
      Logger.log('Error: Data is undefined or null');
      return;
    }
    //Logger.log("DATAAA" + JSON.stringify(data));
  
    // If FLAG is not set or 0, initialize it
    if (flag === null || flag === undefined) {
      flag = 0;
    }
  
    // Iterate over each email using index variable i
    for (var i = flag; i < Object.keys(data).length; i++) {
      //Logger.log("i = " + i + " Max : " + Object.keys(data).length);
      //Logger.log("i = " + i + " Limit : " + limit);
  
      const email = Object.keys(data)[i];
      //Logger.log(Object.keys(data))
      const activitiesByEmail = data[email];
  
      Logger.log("EMAILL : " + i + Object.keys(activitiesByEmail))
  
      // Check if activitiesByEmail is defined
      if (activitiesByEmail && activitiesByEmail[email]) {
        // Reset the flag for each email
        var emailFlag = flag;
  
        // Iterate over each ActivityName
        for (const activityName in activitiesByEmail[email]) {
          const activity = activitiesByEmail[email][activityName];
  
          // Check elapsed time
          if (shouldExit(startTime)) {
            Logger.log('Elapsed time exceeded. Scheduling trigger to resume.');
            writeValueToScriptProperties('FLAG', emailFlag, SCRIPT_PROPERTIES_SERVICE);
            var startTime = new Date().getTime();
            return;
          } else {
            // Increment the emailFlag for each activity
            // Create a row for each activity
            const row = [emailFlag, email, activityName, ...Object.values(activity)];
            newSheet.appendRow(row);
            
          }
        }
        
  
        // Update the main flag for the next email
        //flag = emailFlag;
      } else {
        Logger.log('Error: activitiesByEmail is undefined for email ' + email);
      }
      flag++;
            writeValueToScriptProperties('FLAG', flag, SCRIPT_PROPERTIES_SERVICE);
    }
  
    // Update the FLAG in script properties after processing all emails
    writeValueToScriptProperties('FLAG', flag, SCRIPT_PROPERTIES_SERVICE);
    return;
  }
  
  // Function to check if elapsed time exceeds a threshold
  function shouldExit(startTime) {
    const elapsedMillis = new Date().getTime() - startTime;
    const elapsedMinutes = elapsedMillis / (1000 * 60); // Convert milliseconds to minutes
  
    // Check if elapsed time exceeds the threshold (e.g., 4 minutes)
    return elapsedMinutes > 4;
  }
  
  
  function clearFlag(){
  clearScriptProperty('FLAG');
  }
  
  function initializeFlag(){
    value = 0;
    writeValueToScriptProperties('FLAG',value,SCRIPT_PROPERTIES_SERVICE);
  }