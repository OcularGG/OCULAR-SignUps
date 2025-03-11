/**
 * Composition Data Management
 * 
 * This file contains the structure and data for different party compositions.
 * It provides templates for party setups that can be loaded into the sign-up sheet.
 */

/**
 * Returns the starting row for a composition in the Comp Data sheet
 * 
 * @param {string} compTitle - The title of the composition to look up
 * @return {number} - The starting row for this composition
 */
function getCompStartRow(compTitle) {
  // Map composition titles to their starting rows in the Comp Data sheet
  var compRowMappings = {
    "CLAP KITE": 3,      // A3-A23 and across
    "CLAP SOAK": 26,     // A26-A46 and across
    "CLAP BRAWL": 49,    // A49-A69 and across
    "MONKEY BRAWL": 72,  // A72-A92 and across
    "WALK-IN BRAWL": 95  // A95-A115 and across
  };
  
  return compRowMappings[compTitle] || null;
}

/**
 * Gets composition categories for item counting
 * 
 * @return {Object} - Categories for different item types
 */
function getItemCategories() {
  return {
    pierce: ['Damnation', 'Realmbreaker', 'Lifecurse', 'Incubus', 'Spirithunter'],
    bedrock: ['Bedrock'],
    feyscale: ['Blight', 'Hallowfall', 'Fallen', 'Dawnsong'],
    demon: ['Oathkeepers', 'Bedrock', '1H Arcane', 'Lifecurse', 'Incubus'],
    royalJacket: ['Occult', 'Realmbreaker', 'Spirithunter'],
    knightHelm: ['Hellfire', 'Gauntlets']
  };
}
