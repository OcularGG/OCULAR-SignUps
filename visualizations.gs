/**
 * Composition Visualization System
 * 
 * This file contains functions for generating visualizations and statistics
 * based on party composition data. It handles counting weapons, creating
 * role distribution charts, and providing visual feedback on composition balance.
 */

/**
 * Counts specific items/weapons in the composition and creates visualizations
 * 
 * @param {Object} sheet - The sheet to analyze
 * @return {Object} - Statistics about the composition counts
 */
function countCompItems(sheet) {
  try {
    Logger.log("Starting countCompItems function");
    
    // Define weapon categories for different roles
    const categories = {
      // Weapons that deal armor-piercing damage
      pierce: ['Realmbreaker', 'Lifecurse Staff', 'Spirithunter', 'Damnation Staff', 'Incubus Mace'],
      
      // Special support item
      bedrock: ['Bedrock Mace'],
      
      // Healing weapons
      feyscale: ['Blight Staff', 'Hallowfall', 'Fallen Staff', 'Dawnsong'],
      
      // Support/control items
      demon: ['Oathkeepers', 'Bedrock Mace', '1H Arcane Staff', 'Lifecurse Staff', 'Incubus Mace'],
      
      // Special armor types
      royalJacket: ['Occult Staff', 'Realmbreaker', 'Spirithunter'],
      knightHelm: ['Hellfire Hands', 'Spiked Gauntlets'],
      
      // Role categories for pie chart visualization
      engageTanks: ['Earthrune Staff', 'Hand of Justice', '1H Mace'],
      defensiveTanks: ['Heavy Mace', '1H Hammer', 'Great Arcane Staff', 'Icicle Staff'],
      supports: ['Oathkeepers', 'Bedrock Mace', 'Rootbound Staff', 'Occult Staff', 'Malevolent Locus', '1H Arcane Staff'],
      healers: ['Blight Staff', 'Hallowfall', 'Fallen Staff'],
      battleMounts: ['Chariot', 'Behemoth', 'Battle Eagle', 'Colossus Beetle', 'Siege Ballista', 'Bastion', 'Venom Basilisk'],
      dps: ['Hellfire Hands', 'Rift Glaive', 'Spiked Gauntlets', 'Permafrost Prism', 'Dawnsong']
    };
    
    // Initialize counters for each weapon/role type
    let counts = {
      pierce: 0,
      bedrock: 0,
      feyscale: 0,
      demon: 0,
      royalJacket: 0,
      knightHelm: 0,
      
      // Role counters for pie chart
      engageTank: 0,
      defensiveTank: 0,
      support: 0,
      healer: 0,
      pierceRole: 0, // Different from pierce counter above to avoid naming conflict
      battleMount: 0,
      dps: 0
    };
    
    // Set column headers for counts
    sheet.getRange('F2').setValue("TOTAL PIERCEs");
    sheet.getRange('F3').setValue("TOTAL BEDROCKs");
    sheet.getRange('F4').setValue("TOTAL FEYSCALEs");
    
    sheet.getRange('H2').setValue("TOTAL DEMONs");
    sheet.getRange('H3').setValue("TOTAL ROYAL JACKETs");
    sheet.getRange('H4').setValue("TOTAL KNIGHT HELMs");
    
    // Define all weapon/name range pairs to check
    const rangePairs = [
      // Parties 1-4 (rows 8-26)
      {weapon: 'F8:F26', name: 'G8:G26'},
      {weapon: 'H8:H26', name: 'I8:I26'},
      {weapon: 'J8:J26', name: 'K8:K26'},
      {weapon: 'L8:L26', name: 'M8:M26'},
      
      // Parties 5-8 (rows 28-47)
      {weapon: 'F28:F47', name: 'G28:G47'},
      {weapon: 'H28:H47', name: 'I28:I47'},
      {weapon: 'J28:J47', name: 'K28:K47'},
      {weapon: 'L28:L47', name: 'M28:M47'},
      
      // Parties 9-10 (rows 49-68)
      {weapon: 'F49:F68', name: 'G49:G68'},
      {weapon: 'H49:H68', name: 'I49:I68'}
    ];
    
    // Process each range pair to count weapons
    rangePairs.forEach((pair, idx) => {
      try {
        // Get actual ranges
        const weapons = sheet.getRange(pair.weapon).getValues();
        const names = sheet.getRange(pair.name).getValues();
        
        // Debug log sizes
        Logger.log(`Processing party ${idx+1}: Weapons ${pair.weapon}, Names ${pair.name}`);
        
        // For each row in the range
        for (let i = 0; i < Math.min(weapons.length, names.length); i++) {
          // Get weapon and name values
          const weapon = String(weapons[i][0] || "").trim();
          const name = String(names[i][0] || "").trim();
          
          // Only count if there's both a name and a weapon
          if (name && weapon) {
            // Debug sample data
            if (i < 3) Logger.log(`Found: ${name} with ${weapon}`);
            
            // Count for each category
            if (categories.pierce.includes(weapon)) counts.pierce++;
            if (categories.bedrock.includes(weapon)) counts.bedrock++;
            if (categories.feyscale.includes(weapon)) counts.feyscale++;
            if (categories.demon.includes(weapon)) counts.demon++;
            if (categories.royalJacket.includes(weapon)) counts.royalJacket++;
            if (categories.knightHelm.includes(weapon)) counts.knightHelm++;
            
            // Count for pie chart role categories
            if (categories.engageTanks.includes(weapon)) {
              counts.engageTank++;
            } else if (categories.defensiveTanks.includes(weapon)) {
              counts.defensiveTank++;
            } else if (categories.supports.includes(weapon)) {
              counts.support++;
            } else if (categories.healers.includes(weapon)) {
              counts.healer++;
            } else if (categories.pierce.includes(weapon)) {
              counts.pierceRole++;
            } else if (categories.battleMounts.includes(weapon)) {
              counts.battleMount++;
            } else if (categories.dps.includes(weapon)) {
              counts.dps++;
            }
          }
        }
      } catch (e) {
        Logger.log(`Error processing pair ${pair.weapon}/${pair.name}: ${e}`);
      }
    });
    
    // Set the count values in their respective cells
    sheet.getRange('G2').setValue(counts.pierce);
    sheet.getRange('G3').setValue(counts.bedrock);
    sheet.getRange('G4').setValue(counts.feyscale);
    
    sheet.getRange('I2').setValue(counts.demon);
    sheet.getRange('I3').setValue(counts.royalJacket);
    sheet.getRange('I4').setValue(counts.knightHelm);
    
    // Clear any formatting that might have been applied to these cells
    sheet.getRange('G2:G4').clearFormat();
    sheet.getRange('I2:I4').clearFormat();
    
    // Get colors from Comp Data sheet for consistent role coloring
    try {
      const compSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Comp Data');
      const colorRanges = ['L3', 'L4', 'L5', 'L6', 'L7', 'L8'];
      const roleColors = [];
      
      // Get the background colors from the Comp Data sheet
      for (const cellRef of colorRanges) {
        try {
          roleColors.push(compSheet.getRange(cellRef).getBackground());
        } catch (e) {
          // If there's an error getting colors, add a default one
          roleColors.push(null);
        }
      }
      
      // Create role distribution chart with fetched colors
      createRoleDistributionChartAt(
        sheet, 'N', 1, "OVERALL ROLE DISTRIBUTION",
        counts.engageTank, counts.defensiveTank,
        counts.healer, counts.support,
        counts.pierceRole, counts.battleMount,
        counts.dps, roleColors
      );
    } catch (e) {
      Logger.log(`Error creating role chart: ${e}`);
    }
    
    // Update zerg size based on latest player count
    updateZergSize(sheet);
    
    // Return the counts for use by other functions
    return counts;
  } catch (e) {
    Logger.log("Error in countCompItems: " + e);
    return {};
  }
}

/**
 * Creates visual composition balance indicators
 * Shows bar charts for each weapon type's prevalence
 * 
 * @param {Object} sheet - The sheet to modify
 */
function addCompositionVisualizations(sheet) {
  try {
    // Get total player count for percentage calculations
    const totalPlayers = countTotalPlayers(sheet);
    
    // Create a heading for the balance visualization area
    sheet.getRange('O2').setValue("COMPOSITION BALANCE").setFontWeight('bold');
    
    // Visualize each category with percentage bars
    // Arguments: sheet, cell with count, total players, range for bar, color
    createPercentageBar(sheet, 'G2', totalPlayers, 'P2:R2', '#FF6666'); // Pierce
    createPercentageBar(sheet, 'G3', totalPlayers, 'P3:R3', '#66B2FF'); // Bedrock
    createPercentageBar(sheet, 'G4', totalPlayers, 'P4:R4', '#99FF99'); // Feyscale
    createPercentageBar(sheet, 'I2', totalPlayers, 'P5:R5', '#FF9966'); // Demon
    createPercentageBar(sheet, 'I3', totalPlayers, 'P6:R6', '#CC99FF'); // Royal Jacket
    createPercentageBar(sheet, 'I4', totalPlayers, 'P7:R7', '#FFCC66'); // Knight Helm
    
    // Ensure cells P2-P10 and O10 have white background for clean appearance
    sheet.getRange("P2:P10").setBackground("#FFFFFF");
    sheet.getRange("O10").setBackground("#FFFFFF");
  } catch (e) {
    Logger.log(`Error in addCompositionVisualizations: ${e}`);
  }
}

/**
 * Creates a percentage bar visualization for a specific weapon count
 * 
 * @param {Object} sheet - The sheet to modify
 * @param {string} countCell - Cell containing the count (e.g., 'G2')
 * @param {number} total - Total number of players for percentage calculation
 * @param {string} barRange - Range to use for the bar (e.g., 'P2:R2')
 * @param {string} color - Hex color code for the filled portion of the bar
 */
function createPercentageBar(sheet, countCell, total, barRange, color) {
  try {
    // Get current count from the cell
    const count = sheet.getRange(countCell).getValue() || 0;
    
    // Calculate percentage (avoid division by zero)
    const percentage = (total > 0) ? (count / total) * 100 : 0;
    
    // Reset the bar range to a light gray background
    sheet.getRange(barRange).setBackground("#EEEEEE");
    
    // Calculate which cells to color based on percentage (3 cells total)
    const filledCells = Math.round((percentage / 100) * 3);
    
    // Color the appropriate number of cells
    if (filledCells > 0) {
      // First cell
      sheet.getRange(barRange.split(':')[0]).setBackground(color);
    }
    
    if (filledCells > 1) {
      // Second cell - calculate its address by incrementing the row number
      const secondCell = barRange.split(':')[0].replace(/\d+$/, function(n) { 
        return parseInt(n) + 1; 
      });
      sheet.getRange(secondCell).setBackground(color);
    }
    
    if (filledCells > 2) {
      // Third cell
      sheet.getRange(barRange.split(':')[1]).setBackground(color);
    }
  } catch (e) {
    Logger.log(`Error in createPercentageBar for ${countCell}: ${e}`);
  }
}

/**
 * Checks if the composition balance meets target percentages
 * Note: This function now only does balance checking, not visual formatting
 * 
 * @param {Object} sheet - The sheet to analyze
 */
function checkCompositionBalance(sheet) {
  try {
    // Define target percentages for optimal composition
    const targets = {
      pierce: { min: 15, max: 25 },
      bedrock: { min: 5, max: 10 },
      feyscale: { min: 10, max: 20 },
      demon: { min: 10, max: 20 },
      royalJacket: { min: 5, max: 15 },
      knightHelm: { min: 5, max: 15 }
    };
    
    // Calculate current percentages
    const totalPlayers = countTotalPlayers(sheet);
    if (totalPlayers === 0) return; // Nothing to check if no players
    
    // Calculate percentages for each weapon type
    const piercePercent = (sheet.getRange('G2').getValue() / totalPlayers) * 100;
    const bedrockPercent = (sheet.getRange('G3').getValue() / totalPlayers) * 100;
    const feyscalePercent = (sheet.getRange('G4').getValue() / totalPlayers) * 100;
    const demonPercent = (sheet.getRange('I2').getValue() / totalPlayers) * 100;
    const royalJacketPercent = (sheet.getRange('I3').getValue() / totalPlayers) * 100;
    const knightHelmPercent = (sheet.getRange('I4').getValue() / totalPlayers) * 100;
    
    // Log balance information for debugging
    Logger.log("Composition Balance Check:");
    Logger.log(`Pierce: ${piercePercent.toFixed(1)}% (Target: ${targets.pierce.min}%-${targets.pierce.max}%)`);
    Logger.log(`Bedrock: ${bedrockPercent.toFixed(1)}% (Target: ${targets.bedrock.min}%-${targets.bedrock.max}%)`);
    Logger.log(`Feyscale: ${feyscalePercent.toFixed(1)}% (Target: ${targets.feyscale.min}%-${targets.feyscale.max}%)`);
    
    // Function to check if a value is within target range
    function isBalanced(value, target) {
      return value >= target.min && value <= target.max;
    }
    
    // Return balance status
    return {
      pierce: isBalanced(piercePercent, targets.pierce),
      bedrock: isBalanced(bedrockPercent, targets.bedrock),
      feyscale: isBalanced(feyscalePercent, targets.feyscale),
      demon: isBalanced(demonPercent, targets.demon),
      royalJacket: isBalanced(royalJacketPercent, targets.royalJacket),
      knightHelm: isBalanced(knightHelmPercent, targets.knightHelm)
    };
  } catch (e) {
    Logger.log(`Error in checkCompositionBalance: ${e}`);
    return {};
  }
}

/**
 * Creates an overall role distribution chart showing the breakdown
 * of different roles in the current composition
 * 
 * @param {Object} sheet - The sheet to modify
 */
function createOverallRoleDistribution(sheet) {
  try {
    Logger.log("Starting createOverallRoleDistribution function");
    
    // Clear previous visualization area but preserve formatting
    sheet.getRange('N1:Q16').clearContent();
    
    // Get Comp Data sheet for color reference
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const compSheet = ss.getSheetByName('Comp Data');
    
    // Get colors from Comp Data sheet
    const colorRanges = ['L3', 'L4', 'L5', 'L6', 'L7', 'L8'];
    const roleColors = [];
    
    // Get the background colors from the Comp Data sheet
    for (let i = 0; i < colorRanges.length; i++) {
      try {
        roleColors.push(compSheet.getRange(colorRanges[i]).getBackground());
      } catch (e) {
        roleColors.push(null); // Use null for missing colors
      }
    }
    
    // Define the weapon categories for role detection
    const categories = {
      engageTanks: ['Earthrune Staff', 'Hand of Justice', '1H Mace'],
      defensiveTanks: ['Heavy Mace', '1H Hammer', 'Great Arcane Staff', 'Icicle Staff'],
      healers: ['Blight Staff', 'Hallowfall', 'Fallen Staff'],
      supports: ['Oathkeepers', 'Bedrock Mace', 'Rootbound Staff', 'Occult Staff', 'Malevolent Locus', '1H Arcane Staff'],
      pierces: ['Realmbreaker', 'Lifecurse Staff', 'Spirithunter', 'Damnation Staff', 'Incubus Mace'],
      battleMounts: ['Chariot', 'Behemoth', 'Battle Eagle', 'Colossus Beetle', 'Siege Ballista', 'Bastion', 'Venom Basilisk'],
      dps: ['Hellfire Hands', 'Rift Glaive', 'Spiked Gauntlets', 'Permafrost Prism', 'Dawnsong']
    };
    
    // Counter object for role counts
    const roleCounts = {
      engageTanks: 0,
      defensiveTanks: 0,
      healers: 0,
      supports: 0,
      pierces: 0,
      battleMounts: 0,
      dps: 0
    };
    
    // Define all weapon/name range pairs to check
    const rangePairs = [
      // First section - Parties 1-4
      {weaponRange: 'F8:F26', nameRange: 'G8:G26'},
      {weaponRange: 'H8:H26', nameRange: 'I8:I26'},
      {weaponRange: 'J8:J26', nameRange: 'K8:K26'},
      {weaponRange: 'L8:L26', nameRange: 'M8:M26'},
      
      // Second section - Parties 5-8
      {weaponRange: 'F28:F47', nameRange: 'G28:G47'},
      {weaponRange: 'H28:H47', nameRange: 'I28:I47'},
      {weaponRange: 'J28:J47', nameRange: 'K28:K47'},
      {weaponRange: 'L28:L47', nameRange: 'M28:M47'},
      
      // Third section - Parties 9-10
      {weaponRange: 'F49:F68', nameRange: 'G49:G68'},
      {weaponRange: 'H49:H68', nameRange: 'I49:I68'}
    ];
    
    // Count weapons across all parties
    let processedParties = 0;
    rangePairs.forEach(pair => {
      try {
        const weapons = sheet.getRange(pair.weaponRange).getValues();
        const names = sheet.getRange(pair.nameRange).getValues();
        
        // Log diagnostic info
        const nonEmptyCount = names.filter(n => String(n[0]).trim()).length;
        Logger.log(`Processing party ${++processedParties}: Found ${nonEmptyCount} players`);
        
        // Process each player in this party
        for (let i = 0; i < weapons.length; i++) {
          const weapon = String(weapons[i][0] || "").trim();
          const name = String(names[i][0] || "").trim();
          
          // Only count if both weapon and name are present
          if (name && weapon) {
            // Debug sample data (limited to avoid log spam)
            if (i < 2) Logger.log(`  Found: ${name} with ${weapon}`);
            
            // Check which role this weapon belongs to
            for (const category in categories) {
              if (categories[category].includes(weapon)) {
                roleCounts[category]++;
                break;
              }
            }
          }
        }
      } catch (e) {
        Logger.log(`Error processing ranges ${pair.weaponRange}/${pair.nameRange}: ${e}`);
      }
    });
    
    // Also check "WILL FILL" selections in column E
    try {
      // Retrieve all will-fill values
      const willFillRange = sheet.getRange('E8:E508');
      const willFillValues = willFillRange.getValues();
      
      // Track how many will-fill roles we count
      let willFillCount = 0;
      
      // Process "WILL FILL" selections
      for (let i = 0; i < willFillValues.length; i++) {
        const willFillSelection = String(willFillValues[i][0]).trim();
        
        // Only count if there's a valid role selected
        if (willFillSelection) {
          willFillCount++;
          // Map the WILL FILL selection to the appropriate category
          switch(willFillSelection) {
            case 'DTank':
              roleCounts.defensiveTanks++;
              break;
            case 'Healer':
              roleCounts.healers++;
              break;  
            case 'Support':
              roleCounts.supports++;
              break;
            case 'Pierce':
              roleCounts.pierces++;
              break;
            case 'Bomb':
              roleCounts.dps++;
              break;
            case 'Battle Mount':
              roleCounts.battleMounts++;
              break;
          }
        }
      }
      
      Logger.log(`Counted ${willFillCount} WILL FILL selections`);
    } catch (e) {
      Logger.log(`Error processing WILL FILL selections: ${e}`);
    }
    
    // Create the overall role distribution chart
    createRoleDistributionChartAt(
      sheet, 'N', 1, "OVERALL ROLE DISTRIBUTION",
      roleCounts.engageTanks, roleCounts.defensiveTanks,
      roleCounts.healers, roleCounts.supports,
      roleCounts.pierces, roleCounts.battleMounts,
      roleCounts.dps, roleColors
    );
    
    // Update zerg size counter with latest player count
    updateZergSize(sheet);
  } catch (e) {
    Logger.log(`Error in createOverallRoleDistribution: ${e}`);
  }
}

/**
 * Creates a role distribution chart at the specified location
 * 
 * @param {Object} sheet - The sheet to modify
 * @param {string} column - Starting column letter
 * @param {number} row - Starting row number
 * @param {string} title - Chart title
 * @param {number} engageTankCount - Number of engage tanks
 * @param {number} defensiveTankCount - Number of defensive tanks
 * @param {number} healerCount - Number of healers
 * @param {number} supportCount - Number of supports
 * @param {number} pierceCount - Number of pierce weapons
 * @param {number} battleMountCount - Number of battle mounts
 * @param {number} dpsCount - Number of DPS weapons
 * @param {Array} roleColors - Array of colors for each role
 * @return {number} - Next available row after the chart
 */
function createRoleDistributionChartAt(sheet, column, row, title, 
                                       engageTankCount, defensiveTankCount,
                                       healerCount, supportCount, 
                                       pierceCount, battleMountCount,
                                       dpsCount, roleColors) {
  try {
    // Set chart title with bold formatting
    const titleCell = sheet.getRange(column + row);
    titleCell.setValue(title);
    titleCell.setFontWeight('bold');
    
    // Set up role data
    const labels = ['ENGAGE TANK', 'DEFENSIVE TANK', 'HEALER', 'SUPPORT', 'PIERCE', 'BATTLE MOUNT', 'DPS'];
    const counts = [engageTankCount, defensiveTankCount, healerCount, supportCount, pierceCount, battleMountCount, dpsCount];
    
    // Use colors from Comp Data sheet or fallback to default colors
    let colors;
    if (roleColors && roleColors.length >= 6) {
      // If we have colors from the Comp Data sheet, use those
      colors = roleColors.slice(0, 6);
      // Add a default color for DPS
      colors.push('#FFCC66');
    } else {
      // Fallback colors
      colors = ['#FF6666', '#66B2FF', '#99FF99', '#FF9966', '#CC99FF', '#FFCC66', '#FF9999'];
    }
    
    // Calculate total for percentages
    const total = counts.reduce((a, b) => a + b, 0);
    
    // If no data, display a message
    if (total === 0) {
      sheet.getRange(column + (row + 1)).setValue("No players found");
      return row + 2;
    }
    
    // Set up table headers
    const headerRow = row + 1;
    const dataStartRow = row + 2;
    
    // Set column headers for the role chart
    sheet.getRange(column + headerRow).setValue("ROLE");
    sheet.getRange(String.fromCharCode(column.charCodeAt(0) + 1) + headerRow).setValue("COUNT");
    sheet.getRange(String.fromCharCode(column.charCodeAt(0) + 2) + headerRow).setValue("%");
    
    // Clear any existing content in the distribution column
    const distributionColumnLetter = String.fromCharCode(column.charCodeAt(0) + 3);
    sheet.getRange(`${distributionColumnLetter}${headerRow}:${distributionColumnLetter}${dataStartRow + labels.length}`).clearContent();
    
    // Fill in data rows for each role
    for (let i = 0; i < labels.length; i++) {
      const dataRow = dataStartRow + i;
      const percentage = (counts[i] / total * 100).toFixed(1) + "%";
      
      // Set the role name with colored background
      const roleCell = sheet.getRange(column + dataRow);
      roleCell.setValue(labels[i]);
      roleCell.setBackground(colors[i]);
      roleCell.setFontColor('#000000'); // Black text for readability
      
      // Set the count value
      const countCell = sheet.getRange(String.fromCharCode(column.charCodeAt(0) + 1) + dataRow);
      countCell.setValue(counts[i]);
      countCell.setNumberFormat("0"); // Integer format
      
      // Set the percentage
      sheet.getRange(String.fromCharCode(column.charCodeAt(0) + 2) + dataRow).setValue(percentage);
    }
    
    // Add total row with bold formatting
    const totalRow = dataStartRow + labels.length;
    sheet.getRange(column + totalRow).setValue("TOTAL").setBackground('#EEEEEE').setFontWeight('bold');
    
    // Set total count
    const totalCountCell = sheet.getRange(String.fromCharCode(column.charCodeAt(0) + 1) + totalRow);
    totalCountCell.setValue(total);
    totalCountCell.setFontWeight('bold');
    totalCountCell.setNumberFormat("0"); // Integer format
    
    // Set total percentage (always 100%)
    sheet.getRange(String.fromCharCode(column.charCodeAt(0) + 2) + totalRow)
      .setValue("100%")
      .setFontWeight('bold');
    
    // Ensure visualization area cells have white background
    sheet.getRange("P2:P10").setBackground("#FFFFFF");
    sheet.getRange("O10").setBackground("#FFFFFF");
    
    // Return the next available row
    return totalRow + 2;
  } catch (e) {
    Logger.log(`Error in createRoleDistributionChartAt: ${e}`);
    return row + labels.length + 3; // Return approximate next row
  }
}

/**
 * Counts total players in the Sign-Up sheet for visualization calculations
 * 
 * @param {Object} sheet - The sheet to analyze
 * @return {number} - Total count of players with names
 */
function countTotalPlayers(sheet) {
  try {
    let totalCount = 0;
    
    // Define all name ranges to check
    const nameRanges = [
      'G8:G26', 'I8:I26', 'K8:K26', 'M8:M26', // Parties 1-4
      'G28:G47', 'I28:I47', 'K28:K47', 'M28:M47', // Parties 5-8
      'G49:G68', 'I49:I68' // Parties 9-10
    ];
    
    // Count non-empty names in all ranges
    nameRanges.forEach(range => {
      try {
        const names = sheet.getRange(range).getValues();
        names.forEach(row => {
          if (String(row[0]).trim()) {
            totalCount++;
          }
        });
      } catch (e) {
        Logger.log(`Error counting in range ${range}: ${e}`);
      }
    });
    
    return totalCount;
  } catch (e) {
    Logger.log(`Error in countTotalPlayers: ${e}`);
    return 0;
  }
}

/**
 * Updates the zerg size counter with current player count
 * 
 * @param {Object} sheet - The sheet to update
 */
function updateZergSize(sheet) {
  try {
    const totalPlayers = countTotalPlayers(sheet);
    sheet.getRange('I1').setValue(totalPlayers);
    
    // Add visual indicator based on size
    if (totalPlayers > 40) {
      sheet.getRange('I1').setBackground('#4CAF50'); // Green for good size
    } else if (totalPlayers > 20) {
      sheet.getRange('I1').setBackground('#FFC107'); // Yellow for medium size
    } else {
      sheet.getRange('I1').setBackground('#F44336'); // Red for small size
    }
  } catch (e) {
    Logger.log(`Error in updateZergSize: ${e}`);
  }
}
