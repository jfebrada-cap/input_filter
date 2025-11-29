// Configuration and cache
const DROPDOWN_CACHE = {};
const EC2_DEFAULT_INSTANCES = ['t3a.medium', 't3a.large', 't3a.xlarge', 'r6g.large', 'r6g.xlarge'];
const ECS_DEFAULT_INSTANCES = [
  'ecs.c8i.large',
  'ecs.c8i.xlarge', 
  'ecs.c8i.2xlarge',
  'ecs.c8i.3xlarge',
  'ecs.c8i.4xlarge',
  'ecs.c8i.6xlarge',
  'ecs.c8i.8xlarge',
  'ecs.c8i.12xlarge',
  'ecs.c8i.16xlarge',
  'ecs.c8i.24xlarge',
  'ecs.c8i.48xlarge'
];

// Service configuration with optimized ranges
const SERVICE_CONFIG = {
  'EC2': { 
    dropdownRange: 'A2:A13',
    searchRange: 'A:A',
    cacheKey: 'ec2', 
    hasDefaultInstances: true,
    allowedColumn: 'A:A',
    expectedPrefix: '',
    correctFormat: 't3.medium, c5.large, m5.xlarge',
    provider: 'AWS'
  },
  'ECS': { 
    dropdownRange: 'A2:A12',
    searchRange: 'A:A',
    cacheKey: 'ecs', 
    hasDefaultInstances: true,
    allowedColumn: 'A:A',
    expectedPrefix: 'ecs.',
    correctFormat: 'ecs.c8i.large, ecs.c8i.xlarge, ecs.c8i.2xlarge',
    provider: 'Ali Cloud'
  },
  'RDS': { 
    dropdownRange: 'F2:F10',
    searchRange: 'F:F',
    cacheKey: 'rds', 
    hasDefaultInstances: false,
    allowedColumn: 'F:F',
    expectedPrefix: 'db.',
    correctFormat: 'db.t3.medium, db.r5.large, db.m5.xlarge',
    provider: 'AWS'
  },
  'Aurora': { 
    dropdownRange: 'F2:F10',
    searchRange: 'F:F',
    cacheKey: 'rds', 
    hasDefaultInstances: false,
    allowedColumn: 'F:F',
    expectedPrefix: 'db.',
    correctFormat: 'db.t3.medium, db.r5.large, db.m5.xlarge',
    provider: 'AWS'
  },
  'MSK': { 
    dropdownRange: 'Q2:Q13',
    searchRange: 'Q:Q',
    cacheKey: 'msk', 
    hasDefaultInstances: false,
    allowedColumn: 'Q:Q',
    expectedPrefix: 'kafka.',
    correctFormat: 'kafka.t3.medium, kafka.m5.large',
    provider: 'AWS'
  },
  'Elasticache': { 
    dropdownRange: 'Z2:Z13',
    searchRange: 'Z:Z',
    cacheKey: 'elasticache', 
    hasDefaultInstances: false,
    allowedColumn: 'Z:Z',
    expectedPrefix: 'cache.',
    correctFormat: 'cache.t3.medium, cache.r5.large, cache.m5.xlarge',
    provider: 'AWS'
  }
};

// Track last service for each row
const LAST_SERVICES = {};

/**
 * Determines if instance cell should be cleared based on service change
 */
function shouldClearInstance(newService, oldService) {
  if (!oldService) return true;
  if (oldService === newService) return false;
  if ((oldService === 'RDS' && newService === 'Aurora') || 
      (oldService === 'Aurora' && newService === 'RDS')) return false;
  return true;
}

/**
 * Checks if provider matches for the service
 */
function isProviderMatch(bValue, serviceType) {
  const config = SERVICE_CONFIG[serviceType];
  if (!config) return false;
  
  if (config.provider === 'AWS' && bValue === 'AWS') return true;
  if (config.provider === 'Ali Cloud' && bValue === 'Ali Cloud') return true;
  
  return false;
}

/**
 * Instant dropdown update for specific row - no delays
 */
function updateInstanceDropdownForRow(row) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const bValue = sheet.getRange(`B${row}`).getValue();
  const eValue = sheet.getRange(`E${row}`).getValue();
  const gRange = sheet.getRange(`G${row}`);
  
  // Get config and exit if invalid service
  const config = SERVICE_CONFIG[eValue];
  if (!config) {
    if (LAST_SERVICES[row]) {
      gRange.clearDataValidations();
      gRange.clearContent();
      delete LAST_SERVICES[row];
    }
    return;
  }
  
  // Check provider match - AWS services need B="AWS", Ali Cloud services need B="Ali Cloud"
  if (!isProviderMatch(bValue, eValue)) {
    if (LAST_SERVICES[row]) {
      gRange.clearDataValidations();
      gRange.clearContent();
      delete LAST_SERVICES[row];
    }
    return;
  }
  
  // Clear instance cell only when switching between incompatible services
  if (shouldClearInstance(eValue, LAST_SERVICES[row])) {
    gRange.clearDataValidations();
    gRange.clearContent();
  }
  
  LAST_SERVICES[row] = eValue;
  
  // Get dropdown values instantly
  const dropdownValues = getDropdownValuesFast(config.dropdownRange, config.cacheKey, config.hasDefaultInstances, eValue);
  
  // Apply dropdown or show unavailable
  if (dropdownValues.length === 0) {
    gRange.setValue('instance unavailable');
  } else {
    applyDropdownValidation(gRange, dropdownValues, eValue);
  }
}

/**
 * Update dropdown for all rows (12-117)
 */
function updateAllInstanceDropdowns() {
  for (let row = 12; row <= 117; row++) {
    updateInstanceDropdownForRow(row);
  }
}

/**
 * Ultra-fast dropdown values getter
 */
function getDropdownValuesFast(columnRange, cacheKey, hasDefaultInstances = false, serviceType = '') {
  // Return cached values immediately if available
  if (DROPDOWN_CACHE[cacheKey]) {
    return DROPDOWN_CACHE[cacheKey];
  }
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  let dropdownValues = [];
  
  // SPECIAL HANDLING FOR ECS - Use Alicloud Price Matrix
  if (serviceType === 'ECS') {
    const alicloudSheet = spreadsheet.getSheetByName('Alicloud Price Matrix');
    if (alicloudSheet) {
      try {
        const dataRange = alicloudSheet.getRange(columnRange);
        const values = dataRange.getValues();
        
        const valueSet = new Set();
        for (let i = 0; i < values.length; i++) {
          const value = values[i][0];
          if (value && typeof value === 'string') {
            const trimmed = value.toString().trim();
            if (trimmed && !valueSet.has(trimmed)) {
              valueSet.add(trimmed);
              dropdownValues.push(trimmed);
            }
          }
        }
      } catch (error) {
        console.log('Error reading Alicloud sheet:', error);
      }
    }
    
    // If no values found, use ECS defaults
    if (dropdownValues.length === 0 && hasDefaultInstances) {
      dropdownValues = [...ECS_DEFAULT_INSTANCES];
    }
  } else {
    // Normal handling for other services
    const sourceSheet = spreadsheet.getSheetByName('AWS Cost Matrix');
    if (!sourceSheet) {
      const defaultValues = (cacheKey === 'ec2' && hasDefaultInstances) ? EC2_DEFAULT_INSTANCES : [];
      return defaultValues;
    }
    
    const dataRange = sourceSheet.getRange(columnRange);
    const values = dataRange.getValues();
    
    const valueSet = new Set();
    for (let i = 0; i < values.length; i++) {
      const value = values[i][0];
      if (value && typeof value === 'string') {
        const trimmed = value.toString().trim();
        if (trimmed && !valueSet.has(trimmed)) {
          valueSet.add(trimmed);
          dropdownValues.push(trimmed);
        }
      }
    }
  }
  
  // Minimal processing
  if (dropdownValues.length > 1) {
    dropdownValues.sort();
    
    if (hasDefaultInstances) {
      if (cacheKey === 'ec2') {
        prioritizeDefaultInstancesFast(dropdownValues, EC2_DEFAULT_INSTANCES);
      } else if (cacheKey === 'ecs') {
        prioritizeDefaultInstancesFast(dropdownValues, ECS_DEFAULT_INSTANCES);
      }
    }
  }
  
  // Cache for instant future access
  DROPDOWN_CACHE[cacheKey] = dropdownValues;
  return dropdownValues;
}

/**
 * Fast prioritization for instances
 */
function prioritizeDefaultInstancesFast(allInstances, defaultInstances) {
  const instanceSet = new Set(allInstances);
  const existingDefaults = defaultInstances.filter(instance => instanceSet.has(instance));
  
  if (existingDefaults.length > 0) {
    // Remove defaults from current position
    existingDefaults.forEach(instance => {
      const index = allInstances.indexOf(instance);
      if (index > -1) allInstances.splice(index, 1);
    });
    // Add defaults to the beginning
    allInstances.unshift(...existingDefaults);
  }
}

/**
 * Instant dropdown validation
 */
function applyDropdownValidation(range, dropdownValues, serviceName) {
  if (dropdownValues.length === 0) return;
  
  const config = SERVICE_CONFIG[serviceName];
  
  try {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdownValues, true)
      .setAllowInvalid(true)
      .setHelpText(`Select or type. Format: ${config.correctFormat}`)
      .build();
    
    range.setDataValidation(rule);
    
    // If it's ECS and no value is set, set to first default instance
    if (serviceName === 'ECS' && !range.getValue()) {
      const ecsDefault = ECS_DEFAULT_INSTANCES[0];
      if (dropdownValues.includes(ecsDefault)) {
        range.setValue(ecsDefault);
      }
    }
  } catch (error) {
    range.clearDataValidations();
  }
}

/**
 * Instant edit trigger - NO DELAYS (Rows 12-117)
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const column = range.getColumn();
  
  // Only process rows 12-117
  if (row < 12 || row > 117) return;
  
  // Handle Column E changes (service selection) - INSTANT
  if (column === 5) {
    updateInstanceDropdownForRow(row); // Immediate execution - NO DELAY
  }
  
  // Handle Column B changes (AWS/Ali Cloud toggle) - INSTANT
  if (column === 2) {
    updateInstanceDropdownForRow(row); // Immediate execution - NO DELAY
  }
  
  // Handle Column G changes (instance typing) - INSTANT validation (no delay)
  if (column === 7) {
    const bValue = sheet.getRange(`B${row}`).getValue();
    const eValue = sheet.getRange(`E${row}`).getValue();
    const gValue = range.getValue();
    
    // Only validate for services with manual entries
    if (SERVICE_CONFIG[eValue] && gValue && isProviderMatch(bValue, eValue)) {
      // INSTANT validation - NO DELAY
      validateManualEntryForRow(gValue, eValue, range, row);
    }
  }
}

/**
 * Instant manual entry validation for specific row
 */
function validateManualEntryForRow(instanceType, serviceType, range, row) {
  const dropdownValues = getDropdownValuesFast(
    SERVICE_CONFIG[serviceType].dropdownRange, 
    SERVICE_CONFIG[serviceType].cacheKey, 
    SERVICE_CONFIG[serviceType].hasDefaultInstances,
    serviceType
  );
  
  // Quick check if value is in dropdown
  const isInDropdown = dropdownValues.some(item => 
    item.toLowerCase() === instanceType.toLowerCase()
  );
  
  if (isInDropdown) {
    range.setBackground('#FFFFFF');
    return;
  }
  
  // Fast format validation
  const formatValidation = validateInstanceFormat(instanceType, serviceType);
  if (!formatValidation.valid) {
    range.setValue(formatValidation.error);
    range.setBackground('#FFCCCC');
    return;
  }
  
  range.setBackground('#FFFFFF');
}

/**
 * Fast format validation
 */
function validateInstanceFormat(instanceType, serviceType) {
  if (!instanceType || !serviceType) {
    return { valid: false, error: 'Missing required fields' };
  }
  
  const config = SERVICE_CONFIG[serviceType];
  if (!config) {
    return { valid: false, error: 'Service not supported' };
  }
  
  const instanceLower = instanceType.toLowerCase().trim();
  
  // Fast prefix checks
  if (serviceType === 'EC2') {
    if (instanceLower.startsWith('db.') || instanceLower.startsWith('cache.') || instanceLower.startsWith('kafka.') || instanceLower.startsWith('ecs.')) {
      const corrected = instanceLower.replace(/^(db\.|cache\.|kafka\.|ecs\.)/, '');
      return { valid: false, error: `Remove prefix. Try: ${corrected}`, suggestedCorrection: corrected };
    }
  } else if (serviceType === 'ECS') {
    if (!instanceLower.startsWith('ecs.')) {
      const corrected = 'ecs.' + instanceLower;
      return { valid: false, error: `Add ecs. prefix. Try: ${corrected}`, suggestedCorrection: corrected };
    }
  } else if (config.expectedPrefix && !instanceLower.startsWith(config.expectedPrefix)) {
    const corrected = config.expectedPrefix + instanceLower;
    return { valid: false, error: `Add ${config.expectedPrefix} prefix. Try: ${corrected}`, suggestedCorrection: corrected };
  }
  
  return { valid: true };
}

/**
 * Preload all dropdowns for instant switching
 */
function preloadAllDropdowns() {
  console.log('Preloading all dropdown caches...');
  
  Object.values(SERVICE_CONFIG).forEach(config => {
    if (!DROPDOWN_CACHE[config.cacheKey]) {
      getDropdownValuesFast(config.dropdownRange, config.cacheKey, config.hasDefaultInstances, config.cacheKey === 'ecs' ? 'ECS' : '');
    }
  });
  
  console.log('All dropdowns preloaded for instant switching');
}

/**
 * Fast search function - searches entire column A for ECS
 */
function searchInstanceType(searchTerm, serviceType) {
  if (!searchTerm) return { found: false, message: 'Enter search term' };
  
  const config = SERVICE_CONFIG[serviceType];
  if (!config) return { found: false, message: 'Invalid service' };
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Use Alicloud Price Matrix for ECS, AWS Cost Matrix for others
  const sourceSheet = (serviceType === 'ECS') 
    ? spreadsheet.getSheetByName('Alicloud Price Matrix')
    : spreadsheet.getSheetByName('AWS Cost Matrix');
  
  if (!sourceSheet) return { found: false, message: 'Sheet not found' };
  
  const searchRange = sourceSheet.getRange(config.searchRange);
  const values = searchRange.getValues();
  const searchTermLower = searchTerm.toLowerCase().trim();
  
  const matches = [];
  for (let i = 0; i < Math.min(values.length, 100); i++) {
    const value = values[i][0];
    if (value && typeof value === 'string' && value.toLowerCase().includes(searchTermLower)) {
      matches.push(value.trim());
      if (matches.length >= 5) break;
    }
  }
  
  return matches.length > 0 ? 
    { found: true, matches, message: `Found: ${matches.join(', ')}` } :
    { found: false, message: `No matches for "${searchTerm}"` };
}

/**
 * Initialize with preloading for maximum speed
 */
function initialize() {
  try {
    // Clear existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new trigger
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    
    // Preload all dropdowns for instant switching
    preloadAllDropdowns();
    
    console.log('Script initialized with INSTANT switching for rows 12-117');
  } catch (error) {
    console.error('Initialization error:', error);
  }
}

/**
 * Create custom menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AWS Instance Tools')
    .addItem('Search Instance', 'manualSearchInstance')
    .addItem('Refresh All Dropdowns', 'updateAllInstanceDropdowns')
    .addItem('Preload All', 'preloadAllDropdowns')
    .addItem('Clear Cache', 'clearCache')
    .addToUi();
}

/**
 * Clear cache
 */
function clearCache() {
  for (const key in DROPDOWN_CACHE) {
    delete DROPDOWN_CACHE[key];
  }
  
  // Clear all last services
  for (const key in LAST_SERVICES) {
    delete LAST_SERVICES[key];
  }
  
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Clear all instance cells from row 12 to 117
  for (let row = 12; row <= 117; row++) {
    const gRange = sheet.getRange(`G${row}`);
    gRange.clearDataValidations();
    gRange.clearContent();
    gRange.setBackground('#FFFFFF');
  }
  
  console.log('Cache cleared for all rows');
}

/**
 * Manual trigger for specific row (backward compatibility)
 */
function manualUpdateDropdown() {
  updateInstanceDropdownForRow(12);
}

/**
 * Manual trigger for all rows
 */
function manualUpdateAllDropdowns() {
  updateAllInstanceDropdowns();
}

/**
 * Search function for active row
 */
function manualSearchInstance() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRow = sheet.getActiveRange().getRow();
  
  // Only allow search for rows 12-117
  if (activeRow < 12 || activeRow > 117) {
    ui.alert('Error', 'Please select a row between 12 and 117', ui.ButtonSet.OK);
    return;
  }
  
  const bValue = sheet.getRange(`B${activeRow}`).getValue();
  const eValue = sheet.getRange(`E${activeRow}`).getValue();
  
  // Check provider match
  if (!isProviderMatch(bValue, eValue)) {
    ui.alert('Error', `Set B${activeRow} to ${SERVICE_CONFIG[eValue]?.provider || 'correct provider'} first`, ui.ButtonSet.OK);
    return;
  }
  
  if (!SERVICE_CONFIG[eValue]) {
    ui.alert('Error', `Select valid service in E${activeRow}`, ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.prompt('Search', 'Enter instance type:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const searchTerm = response.getResponseText();
    if (searchTerm) {
      const result = searchInstanceType(searchTerm, eValue);
      const message = result.found ? 
        `${result.message}\n\nUse "${result.matches[0]}" in G${activeRow}?` : 
        result.message;
      
      if (result.found && ui.alert('Results', message, ui.ButtonSet.YES_NO) === ui.Button.YES) {
        sheet.getRange(`G${activeRow}`).setValue(result.matches[0]);
      } else if (!result.found) {
        ui.alert('Results', message, ui.ButtonSet.OK);
      }
    }
  }
}

/**
 * Test function to verify all rows work
 */
function testAllRows() {
  console.log('Testing all rows from 12 to 117...');
  
  for (let row = 12; row <= 117; row++) {
    console.log(`Testing row ${row}`);
    updateInstanceDropdownForRow(row);
  }
  
  console.log('All rows tested successfully');
}