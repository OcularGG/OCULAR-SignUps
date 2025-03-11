# Event Scheduler AppScript

A Google Sheets-based tool for organizing and scheduling in-game events, particularly aimed at managing OCULAR alliance player compositions for Albion Online events.

## 📋 Overview

The Event Scheduler is a Google Sheets AppScript application that helps event organizers:
- Schedule events with details for massing and zoning times
- Assign callers and roles
- Load pre-defined party compositions
- Track player participation and readiness
- Visualize team balance and composition statistics

## 🔧 Features

- **Event Management**: Create and manage game events with detailed timing information
- **Composition Templates**: Load pre-defined party compositions for different strategies
- **Role Visualization**: Automatic counting and visualization of team roles and balance
- **Player Tracking**: Track player readiness with checkbox confirmation
- **Weapon Distribution**: Monitor critical weapon counts for balanced gameplay

## 📁 Files in this Repository

### AppScript Files
- **code.gs**: Core functionality for event data handling
- **main.gs**: Menu creation and entry points for the application
- **compositions.gs**: Party composition data and template definitions
- **setup.gs**: Sheet initialization and structure setup
- **visualizations.gs**: Data visualization and statistics generation

### HTML Files
- **page.html**: User interface for event scheduling sidebar

### CSV Template Files
- **Sign-Up.csv**: Template for the main Sign-Up sheet
- **Comp-Data.csv**: Template for composition data and weapon configurations

## 🚀 Setup Instructions

### Step 1: Create a new Google Sheet
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Give it a meaningful name (e.g., "OCULAR Event Scheduler")

### Step 2: Import CSV Files
1. Import the Sign-Up.csv file:
   - Go to File > Import
   - Upload the Sign-Up.csv file
   - Select "Replace current sheet" option
   - Click "Import data"
   - Rename the sheet to exactly "Sign-Up"

2. Import the Comp-Data.csv file:
   - Go to File > Import
   - Upload the Comp-Data.csv file
   - Select "Insert new sheet(s)" option
   - Click "Import data"
   - Ensure the sheet is named exactly "Comp Data"

### Step 3: Set Up AppScript
1. Open the Script Editor:
   - Go to Extensions > Apps Script
   - This will open the Google Apps Script editor in a new tab

2. Create the script files:
   - Delete any default code (like `function myFunction() {}`)
   - For each .gs file in the repository:
     - Click the + icon next to "Files" to create a new file
     - Name the file exactly as it appears in the repository (e.g., "code.gs")
     - Paste the code from the repository file
     - Save (Ctrl+S or ⌘+S)

3. Create the HTML file:
   - Click the + icon next to "Files"
   - Select "HTML" from the dropdown menu
   - Name the file "Page" (without the .html extension)
   - Paste the HTML code from page.html
   - Save (Ctrl+S or ⌘+S)

### Step 4: Run Initial Setup
1. Go back to your Google Sheet
2. Refresh the page
3. You should now see an "Event Scheduler" menu in the top navigation
4. If the "Event Scheduler" menu doesn't appear, run the initialization:
   - Return to the Apps Script editor
   - Select the function `onOpen` from the dropdown menu
   - Click the "Run" button (▶️)
   - Grant any permissions requested

### Step 5: Setup Sheet Structure (First Time Only)
1. From the Google Sheet, click "Advanced Tools" > "Fix Range Issues"
2. This will ensure all ranges are properly configured
3. You may need to grant additional permissions

## 📝 Using the Event Scheduler

### Creating a New Event
1. From your Google Sheet, click "Event Scheduler" > "Open Event Scheduler"
2. Fill in the event details:
   - Select Event Type
   - Set Massing Date/Time
   - Set Zoning Date/Time
   - Select Main, Secondary, and Escape Callers
   - Choose a Party Composition (optional)
3. Click "Submit" to save the event and load the selected composition

### Managing Player Sign-ups
1. Players can enter their names in the NAME columns
2. Weapons will be pre-filled based on the selected composition
3. When players confirm their attendance, check the "READY" box
4. Statistics and visualizations will update automatically

## 🛠 Customizing the Scheduler

### Modifying Compositions
1. Edit the "Comp Data" sheet to modify existing compositions or add new ones
2. Follow the existing structure where:
   - Each composition starts with its name (e.g., "CLAP KITE")
   - Columns A-D represent first set of parties
   - Columns E-H represent second set of parties
   - Columns I-J represent third set of parties

### Adding New Weapon Types
1. Locate the appropriate arrays in compositions.gs or visualizations.gs
2. Add new weapon names to the appropriate category arrays
3. For visualization purposes, you may need to assign colors in the "Comp Data" sheet

## 🤝 Contributing

Contributions are welcome! If you have improvements or bug fixes:
1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to your branch
5. Create a new Pull Request

## ⚠️ Troubleshooting

- **"Script Error: Sheet not found"**: Ensure your sheets are named exactly "Sign-Up" and "Comp Data" (case sensitive)
- **Missing menu items**: Try refreshing the page or running the onOpen function manually
- **Permission errors**: Make sure to grant all required permissions when prompted

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgements

- OCULAR alliance for the inspiration and use case
- Contributors who have helped refine the tool
