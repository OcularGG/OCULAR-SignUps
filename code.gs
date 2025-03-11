function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Event Scheduler')
    .addItem('Open Event Scheduler', 'showSidebar')
    .addToUi();
  
  // Add advanced tools menu (from main.gs)
  createAdvancedToolsMenu(SpreadsheetApp.getUi());
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
    .setTitle('Event Scheduler')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Add error handling to setEventData
function setEventData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Sign-Up');
    
    if (!sheet) {
      // If Sign-Up sheet doesn't exist, try to create it
      sheet = initializeSheet();
    }
    
    // Set content type in D1
    if (data.contentType) {
      sheet.getRange('D1').setValue(data.contentType);
    }
    
    // Set date/times
    if (data.massingDateTime) {
      var massingDate = new Date(data.massingDateTime);
      var formattedMassingDateTime = formatDateTime(massingDate);
      sheet.getRange('D2').setValue(formattedMassingDateTime);
    }
    
    if (data.zoningDateTime) {
      var zoningDate = new Date(data.zoningDateTime);
      
      // For zoning time, only show the time in the format "19:30 UTC"
      var hours = zoningDate.getUTCHours().toString().padStart(2, '0');
      var minutes = zoningDate.getUTCMinutes().toString().padStart(2, '0');
      var formattedZoningTime = hours + ":" + minutes + " UTC";
      
      sheet.getRange('D3').setValue(formattedZoningTime);
    }
    
    // Set callers
    if (data.caller) {
      sheet.getRange('D4').setValue(data.caller);
    }
    
    if (data.secondaryCaller) {
      sheet.getRange('D5').setValue(data.secondaryCaller);
    }
    
    if (data.escapeCaller) {
      sheet.getRange('D6').setValue(data.escapeCaller);
    }
    
    // Handle comp data if selected
    if (data.compTitle) {
      loadCompData(data.compTitle);
    }
    
    return true;
  } catch (e) {
    Logger.log("Error in setEventData: " + e.toString());
    throw e; // Re-throw to be handled by the client
  }
}

function loadCompData(compTitle) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var compSheet = ss.getSheetByName('Comp Data');
  var signUpSheet = ss.getSheetByName('Sign-Up');
  
  // Check if both sheets exist
  if (!compSheet) {
    throw new Error("Comp Data sheet not found! Please create this sheet first.");
  }
  if (!signUpSheet) {
    throw new Error("Sign-Up sheet not found! Please run initializeSheet() first.");
  }
  
  // Get the starting row for this comp using the helper function
  var startRow = getCompStartRow(compTitle);
  
  if (!startRow) {
    throw new Error("Composition not recognized: " + compTitle);
  }
  
  // First, clear existing data in the destination ranges
  clearExistingCompData(signUpSheet);
  
  // Map source columns to destination ranges based on the new organization
  // First section: columns A-D from source to E-K in target (with gaps for names)
  var firstSectionColumnMappings = [
    { source: "A", dest: "E" }, // Column A to E (leaving F empty for names)
    { source: "B", dest: "G" }, // Column B to G (leaving H empty for names)
    { source: "C", dest: "I" }, // Column C to I (leaving J empty for names)
    { source: "D", dest: "K" }  // Column D to K (leaving L empty for names)
  ];
  
  // Second section: columns E-H from source to E28-K48 in target (with gaps for names)
  var secondSectionColumnMappings = [
    { source: "E", dest: "E", startRow: 28 }, // Column E to E28 (leaving F empty for names)
    { source: "F", dest: "G", startRow: 28 }, // Column F to G28 (leaving H empty for names)
    { source: "G", dest: "I", startRow: 28 }, // Column G to I28 (leaving J empty for names)
    { source: "H", dest: "K", startRow: 28 }  // Column H to K28 (leaving L empty for names)
  ];
  
  // Third section: columns I-J from source to E50-G70 in target (with gaps for names)
  var thirdSectionColumnMappings = [
    { source: "I", dest: "E", startRow: 50 }, // Column I to E50 (leaving F empty for names)
    { source: "J", dest: "G", startRow: 50 }  // Column J to G50 (leaving H empty for names)
  ];
  
  // Prepare arrays to store all cell values for counting and unique weapon names
  var allValues = [];
  var allWeapons = new Set();
  
  // Loop through first section mappings and copy the data
  for (var i = 0; i < firstSectionColumnMappings.length; i++) {
    var mapping = firstSectionColumnMappings[i];
    
    // Get the source range (20 rows for 20 players)
    var sourceRange = compSheet.getRange(mapping.source + startRow + ":" + mapping.source + (startRow + 19));
    var sourceData = sourceRange.getValues();
    allValues = allValues.concat(sourceData);
    
    // Extract weapon names from source data
    for (var j = 0; j < sourceData.length; j++) {
      var cellValue = String(sourceData[j][0]).trim();
      if (cellValue) {
        allWeapons.add(cellValue);
      }
    }
    
    // Get the destination range (starting from row 7)
    var destRange = signUpSheet.getRange(mapping.dest + "7:" + mapping.dest + "26");
    destRange.setValues(sourceData);
  }
  
  // Loop through second section mappings and copy the data
  for (var i = 0; i < secondSectionColumnMappings.length; i++) {
    var mapping = secondSectionColumnMappings[i];
    
    // Get the source range (20 rows for 20 players)
    var sourceRange = compSheet.getRange(mapping.source + startRow + ":" + mapping.source + (startRow + 19));
    var sourceData = sourceRange.getValues();
    allValues = allValues.concat(sourceData);
    
    // Extract weapon names from source data
    for (var j = 0; j < sourceData.length; j++) {
      var cellValue = String(sourceData[j][0]).trim();
      if (cellValue) {
        allWeapons.add(cellValue);
      }
    }
    
    // Get the destination range (starting from row 28)
    var destRange = signUpSheet.getRange(mapping.dest + mapping.startRow + ":" + mapping.dest + (mapping.startRow + 19));
    destRange.setValues(sourceData);
  }
  
  // Loop through third section mappings and copy the data
  for (var i = 0; i < thirdSectionColumnMappings.length; i++) {
    var mapping = thirdSectionColumnMappings[i];
    
    // Get the source range (20 rows for 20 players)
    var sourceRange = compSheet.getRange(mapping.source + startRow + ":" + mapping.source + (startRow + 19));
    var sourceData = sourceRange.getValues();
    allValues = allValues.concat(sourceData);
    
    // Extract weapon names from source data
    for (var j = 0; j < sourceData.length; j++) {
      var cellValue = String(sourceData[j][0]).trim();
      if (cellValue) {
        allWeapons.add(cellValue);
      }
    }
    
    // Get the destination range (starting from row 50)
    var destRange = signUpSheet.getRange(mapping.dest + mapping.startRow + ":" + mapping.dest + (mapping.startRow + 19));
    destRange.setValues(sourceData);
  }
  
  // Add section headers - REMOVED setting "Parties 1-4" in E5
  signUpSheet.getRange('E27').setValue("Parties 5-8");
  signUpSheet.getRange('E49').setValue("Parties 9-10");
  
  // Set the composition title in cell F1
  signUpSheet.getRange('F1').setValue(compTitle);
  
  // Count specific items in the composition - with updated code to check for names
  countCompItems(signUpSheet);
  
  // Create dropdowns with unique weapon names
  createWeaponDropdowns(Array.from(allWeapons).sort(), signUpSheet);
}

function createWeaponDropdowns(uniqueWeapons, sheet) {
  // Clear existing dropdowns
  sheet.getRange('B8:D25').clearContent().clearDataValidations();
  
  // Skip if no weapons found
  if (uniqueWeapons.length === 0) return;
  
  // Create data validation rule
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueWeapons, true)
    .setAllowInvalid(false)
    .build();
  
  // Add headers
  sheet.getRange('B7').setValue("Reserved Weapons");
  sheet.getRange('C7').setValue("Reserved Weapons");
  sheet.getRange('D7').setValue("Reserved Weapons");
  
  // Apply rule to columns B, C, and D (rows 8-25)
  sheet.getRange('B8:B25').setDataValidation(rule);
  sheet.getRange('C8:C25').setDataValidation(rule);
  sheet.getRange('D8:D25').setDataValidation(rule);
}

function countCompItems(sheet) {
  // Define the items to count
  var pierceWeapons = ['Damnation', 'Realmbreaker', 'Lifecurse', 'Incubus', 'Spirithunter'];
  var bedrockItems = ['Bedrock'];
  var feyscaleWeapons = ['Blight', 'Hallowfall', 'Fallen', 'Dawnsong'];
  
  // Initialize counters
  var pierceCount = 0;
  var bedrockCount = 0;
  var feyscaleCount = 0;
  
  // Update the labels with the new wording
  sheet.getRange('E2').setValue("TOTAL PIERCEs");
  sheet.getRange('E3').setValue("TOTAL BEDROCKs");
  sheet.getRange('E4').setValue("TOTAL FEYSCALEs");
  
  // First section (Parties 1-4): E7:L26
  countItemsInSection('E7:E26', 'F7:F26');
  countItemsInSection('G7:G26', 'H7:H26');
  countItemsInSection('I7:I26', 'J7:J26');
  countItemsInSection('K7:K26', 'L7:L26');
  
  // Second section (Parties 5-8): E28:L47
  countItemsInSection('E28:E47', 'F28:F47');
  countItemsInSection('G28:G47', 'H28:H47');
  countItemsInSection('I28:I47', 'J28:J47');
  countItemsInSection('K28:K47', 'L28:L47');
  
  // Third section (Parties 9-10): E50:H69
  countItemsInSection('E50:E69', 'F50:F69');
  countItemsInSection('G50:G69', 'H50:H69');
  
  // Helper function to count items in a section - only if there's a name present
  function countItemsInSection(weaponRange, nameRange) {
    var weapons = sheet.getRange(weaponRange).getValues();
    var names = sheet.getRange(nameRange).getValues();
    
    for (var i = 0; i < weapons.length; i++) {
      var weapon = String(weapons[i][0]).trim();
      var name = String(names[i][0]).trim();
      
      // Only count if there's a name present in the adjacent cell
      if (name && weapon) {
        if (pierceWeapons.includes(weapon)) {
          pierceCount++;
        } else if (bedrockItems.includes(weapon)) {
          bedrockCount++;
        } else if (feyscaleWeapons.includes(weapon)) {
          feyscaleCount++;
        }
      }
    }
  }
  
  // Set the count values
  sheet.getRange('F2').setValue(pierceCount);
  sheet.getRange('F3').setValue(bedrockCount);
  sheet.getRange('F4').setValue(feyscaleCount);
}

function formatDateTime(date) {
  // Format the date as "Saturday, March 9th @ 19:00 UTC"
  var options = { weekday: 'long', month: 'long', day: 'numeric', timeZone: 'UTC' };
  var formattedDate = date.toLocaleDateString('en-US', options);
  
  // Add the ordinal suffix to the day
  var day = date.getUTCDate();
  var suffix = getDaySuffix(day);
  formattedDate = formattedDate.replace(day, day + suffix);
  
  // Add the time in 24-hour format
  var hours = date.getUTCHours().toString().padStart(2, '0');
  var minutes = date.getUTCMinutes().toString().padStart(2, '0');
  var formattedDateTime = formattedDate + " @ " + hours + ":" + minutes + " UTC";
  
  return formattedDateTime;
}

function getDaySuffix(day) {
  if (day > 3 && day < 21) return 'th';
  switch (day % 10) {
    case 1: return 'st';
    case 2: return 'nd';
    case 3: return 'rd';
    default: return 'th';
  }
}

function clearExistingCompData(sheet) {
  // Clear first section (rows 7-26)
  sheet.getRange('E7:L26').clearContent();
  
  // Clear second section (rows 28-47)
  sheet.getRange('E28:L47').clearContent();
  
  // Clear third section (rows 50-69)
  sheet.getRange('E50:H69').clearContent();
  
  // Clear section headers
  sheet.getRange('E5').clearContent();
  sheet.getRange('E27').clearContent();
  sheet.getRange('E49').clearContent();
  
  // Clear the count values
  sheet.getRange('F2:F4').clearContent();
  
  // Clear weapon dropdowns
  sheet.getRange('B7:D25').clearContent().clearDataValidations();
  
  // Clear composition title
  sheet.getRange('F1').clearContent();
}