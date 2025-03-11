/**
 * Sheet Setup and Initialization Utilities
 * 
 * This file contains functions for setting up and maintaining the sheet structure,
 * including initializing essential components like checkboxes, dropdowns, and formatting.
 * These functions ensure the sheet is ready for use with all required elements.
 */

/**
 * Main initialization function that sets up a new or existing sheet
 * with all required elements for the Event Scheduler
 * 
 * @return {Object} - The initialized sheet object
 */
function initializeSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sign-Up');
  
  // Create Sign-Up sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Sign-Up');
    Logger.log("Created new Sign-Up sheet");
  }
  
  // Configure sheet with required elements
  setupBasicSheetStructure(sheet);
  addCheckboxesColumn(sheet);
  setupHeaders(sheet);
  setupDropdowns(sheet);
  formatSheet(sheet);
  
  // Show confirmation message
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Sheet initialized successfully. Ready to use Event Scheduler.",
    "Setup Complete",
    5
  );
  
  return sheet;
}

/**
 * Sets up the basic sheet structure with proper labels
 * and column/row organization
 * 
 * @param {Object} sheet - The sheet to initialize
 */
function setupBasicSheetStructure(sheet) {
  // Set up title area (E1-I1)
  sheet.getRange('E1').setValue("Content Type:").setFontWeight('bold');
  sheet.getRange('E2').setValue("Massing Time:").setFontWeight('bold');
  sheet.getRange('E3').setValue("Zoning Time:").setFontWeight('bold');
  sheet.getRange('E4').setValue("Main Caller:").setFontWeight('bold');
  sheet.getRange('E5').setValue("Secondary Caller:").setFontWeight('bold');
  sheet.getRange('E6').setValue("Escape Caller:").setFontWeight('bold');
  
  // Prepare visualization area
  sheet.getRange('N1').setValue("VISUALIZATIONS").setFontWeight('bold');
  
  // Initialize ZERG SIZE area
  sheet.getRange('H1').setValue("ZERG SIZE");
  sheet.getRange('I1').setValue("0");
  
  // Freeze rows 1-7 to keep headers visible when scrolling
  sheet.setFrozenRows(7);
}

/**
 * Adds "Ready" checkboxes to column A and sets up confirmation indicator in column B
 * 
 * @param {Object} sheet - The sheet to modify
 */
function addCheckboxesColumn(sheet) {
  // Create checkbox data validation rule
  var checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  
  // Apply checkboxes to range A8:A508
  sheet.getRange('A8:A508').setDataValidation(checkboxRule);
  
  // Set the header for column A and add helpful note
  var readyHeader = sheet.getRange('A7');
  readyHeader.setValue("READY");
  readyHeader.setFontWeight('bold');
  readyHeader.setNote('Check the box when a player has confirmed they are ready');
  
  // Set header for column B with checkmark symbol and note explaining its purpose
  var checkmarkHeader = sheet.getRange('B7');
  checkmarkHeader.setValue("✓");
  checkmarkHeader.setFontWeight('bold');
  checkmarkHeader.setHorizontalAlignment('center');
  checkmarkHeader.setNote('This cell will turn green when the checkbox is checked');
  
  // Set column widths for better usability
  sheet.setColumnWidth(1, 50);  // Column A - narrow for checkboxes
  sheet.setColumnWidth(2, 60);  // Column B - confirmation indicator
}

/**
 * Sets up all column headers with proper formatting
 * 
 * @param {Object} sheet - The sheet to modify
 */
function setupHeaders(sheet) {
  // Define all headers
  var headers = [
    { col: 'C', value: "PREFERRED", width: 130 },
    { col: 'D', value: "SECONDARY", width: 130 },
    { col: 'E', value: "WILL FILL", width: 130 },
    { col: 'F', value: "WEAPON", width: 130 },
    { col: 'G', value: "NAME", width: 150 },
    { col: 'H', value: "WEAPON", width: 130 },
    { col: 'I', value: "NAME", width: 150 },
    { col: 'J', value: "WEAPON", width: 130 },
    { col: 'K', value: "NAME", width: 150 },
    { col: 'L', value: "WEAPON", width: 130 },
    { col: 'M', value: "NAME", width: 150 }
  ];
  
  // Apply headers and set column widths
  headers.forEach(header => {
    var cell = sheet.getRange(header.col + '7');
    cell.setValue(header.value);
    sheet.setColumnWidth(columnToIndex(header.col), header.width);
  });
  
  // Format all headers with blue background and white text
  var headerRange = sheet.getRange('A7:M7');
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#0F52BA');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');
}

/**
 * Sets up dropdown menus for weapon selection and role fields
 * 
 * @param {Object} sheet - The sheet to modify
 */
function setupDropdowns(sheet) {
  // Create role categories dropdown for "Will Fill" column
  var roleCategories = ['DTank', 'Healer', 'Support', 'Pierce', 'Bomb', 'Battle Mount'];
  var roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(roleCategories, true)
    .setAllowInvalid(false)
    .build();
  
  // Apply role categories rule to WILL FILL dropdown (column E)
  sheet.getRange('E8:E508').setDataValidation(roleRule);
  
  // Add section labels to B column to organize the parties
  sheet.getRange('B8').setValue("PARTY 1-4").setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B28').setValue("PARTY 5-8").setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B49').setValue("PARTY 9-10").setFontWeight('bold').setHorizontalAlignment('center');
  
  // Set up explanatory text for the "Will Fill" dropdown
  enhanceWillFillDropdown(sheet);
}

/**
 * Enhances the "Will Fill" dropdown with helpful notes and visual cues
 * 
 * @param {Object} sheet - The sheet to modify
 */
function enhanceWillFillDropdown(sheet) {
  // Add a custom note explaining the dropdown purpose
  var willFillNote = "Select the role you're willing to fill from the dropdown list.";
  
  // Add blue border to highlight the "Will Fill" column
  sheet.getRange('E8:E508').setBorder(true, true, true, true, null, null, '#4285F4', SpreadsheetApp.BorderStyle.SOLID);
  
  // Add color-coded legend for role categories at the bottom of the sheet
  var legendRange = sheet.getRange('E509:E515');
  legendRange.clearContent();
  
  // Set up the legend with role categories and descriptions
  var legend = [
    ["Role Categories:"],
    ["DTank: Defensive Tank (Heavy Mace, 1H Hammer, etc.)"],
    ["Healer: Any healing weapon (Blight, Hallowfall, etc.)"],
    ["Support: Utility support (Bedrock, Oathkeepers, etc.)"],
    ["Pierce: Armor-piercing weapons (Spirithunter, Incubus, etc.)"],
    ["Bomb: AoE damage dealers (Hellfire, Gauntlets, etc.)"],
    ["Battle Mount: Special mounts (Chariot, Behemoth, etc.)"]
  ];
  
  legendRange.setValues(legend);
  legendRange.setFontStyle("italic");
  legendRange.setFontSize(9);
  sheet.getRange('E509').setFontWeight("bold");
}

/**
 * Formats the sheet for better visual organization and readability
 * 
 * @param {Object} sheet - The sheet to format
 */
function formatSheet(sheet) {
  // Add alternating colors to rows for better readability
  for (var i = 8; i <= 508; i++) {
    var rowColor = (i % 2 === 0) ? '#f3f3f3' : '#ffffff';
    sheet.getRange(i, 3, 1, 11).setBackground(rowColor); // Columns C through M
  }
  
  // Add visual separators between party sections
  sheet.getRange('A27:M27').setBackground('#e6e6e6').setBorder(true, true, true, true, null, null);
  sheet.getRange('A48:M48').setBackground('#e6e6e6').setBorder(true, true, true, true, null, null);
  
  // Add borders to main section areas
  sheet.getRange('A7:M508').setBorder(true, true, true, true, true, true);
  
  // Format the title area
  sheet.getRange('E1:G6').setBorder(true, true, true, true, true, true);
  
  // Make the player name areas stand out with light blue background
  var nameRanges = [
    'G8:G26', 'I8:I26', 'K8:K26', 'M8:M26',
    'G28:G47', 'I28:I47', 'K28:K47', 'M28:M47',
    'G49:G68', 'I49:I68'
  ];
  
  // Apply a subtle highlight to name columns
  nameRanges.forEach(range => {
    sheet.getRange(range).setBackground('#f0f8ff');
  });
}

/**
 * Reset all checkboxes in column A and their indicators in column B
 * Used to start a new session with fresh status indicators
 */
function resetCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sign-Up');
  
  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Sign-Up sheet not found! Please run the initializeSheet function first.",
      "Error",
      5
    );
    return;
  }
  
  // Reset all checkboxes to unchecked
  sheet.getRange('A8:A508').setValue(false);
  
  // Reset all B column cells to white background
  sheet.getRange('B8:B508').setBackground('#FFFFFF');
  
  // Confirm reset with toast notification
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'All ready status indicators have been reset.',
    'Reset Complete',
    3
  );
}

/**
 * Helper function to convert column letter to index
 * 
 * @param {string} column - Column letter (A, B, C, etc.)
 * @return {number} - Column index (1-based)
 */
function columnToIndex(column) {
  var result = 0;
  for (var i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result;
}
