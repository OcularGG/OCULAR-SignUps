/**
 * Event Scheduler - Core System Entry Points
 * 
 * This file contains the main entry points for the Event Scheduler application,
 * including menu setup, sidebar display, and core function connections.
 * It serves as the central hub that bridges the various components together.
 */

// ISSUE: Duplicate onOpen function (already exists in code.gs)
// Fix: Rename this to avoid conflicts, or merge functionality with the other one
function initializeApp() {
  var ui = SpreadsheetApp.getUi();
  
  // Create the main menu with essential options
  ui.createMenu('Event Scheduler')
    .addItem('Open Event Scheduler', 'showSidebar')
    .addItem('Review Similar Names', 'showNameReviewDialog')
    .addToUi();
  
  // Create an advanced tools menu
  createAdvancedToolsMenu(ui);
}

/**
 * Creates the Advanced Tools submenu
 * This is implemented as a separate function to improve readability
 * 
 * @param {Object} ui - SpreadsheetApp UI object
 */
function createAdvancedToolsMenu(ui) {
  try {
    var advancedMenu = ui.createMenu('Advanced Tools')
      .addItem('Run Full Diagnostics', 'runDiagnostics')
      .addItem('Test Player Count', 'testPlayerCount')
      .addItem('Check for Duplicate Names', 'checkDuplicateNames')
      .addItem('Fix Range Issues', 'fixRangeIssues')
      .addItem('Reset Ready Checkboxes', 'resetCheckboxes');
    
    // Add as standalone menu if needed
    advancedMenu.addToUi();
  } catch (e) {
    Logger.log("Error creating advanced tools menu: " + e.toString());
  }
}

/**
 * Opens the event scheduler sidebar
 * Creates and displays the HTML UI from the Page.html file
 */
function showSidebar() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('Event Scheduler')
      .setWidth(300);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    Logger.log("Error showing sidebar: " + e.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Error opening scheduler: " + e.message,
      "Error",
      5
    );
  }
}

/**
 * Processes composition selection from the UI
 * Handles loading the selected composition template
 * 
 * @param {string} compTitle - The name of the composition to load
 * @return {Object} - Result object with success flag and message
 */
function processCompSelection(compTitle) {
  try {
    // Validate input
    if (!compTitle) {
      return { 
        success: false, 
        message: "Error: No composition selected" 
      };
    }
    
    // Load the composition data
    loadCompData(compTitle);
    
    // Return success response
    return { 
      success: true, 
      message: "Loaded " + compTitle + " composition" 
    };
  } catch (e) {
    // Log the error and return failure response
    Logger.log("Error loading composition: " + e.toString());
    return { 
      success: false, 
      message: "Error: " + e.toString() 
    };
  }
}

/**
 * Shows the name review dialog for identifying similar player names
 * Helps maintain consistent naming across the sheet
 */
function showNameReviewDialog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('NameReview')
      .setWidth(600)
      .setHeight(400)
      .setTitle('Review Similar Names');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Review Similar Names');
  } catch (e) {
    Logger.log("Error showing name review dialog: " + e.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Error opening name review: " + e.message,
      "Error",
      5
    );
  }
}

/**
 * Utility function to display a toast message
 * Centralizes toast creation for consistent UI
 * 
 * @param {string} message - Message to show
 * @param {string} title - Title of the toast
 * @param {number} duration - Duration in seconds
 */
function showMessage(message, title, duration) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title || "Info", duration || 5);
}
