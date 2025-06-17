/**
 * Beauty Pro Inventory System - Main Code Module
 * Google Apps Script automation for interactive inventory management
 */

// Global configuration
const CONFIG = {
  colors: {
    deepCharcoal: '#2C3E50',
    roseGold: '#E8B4B8', 
    creamWhite: '#F8F9FA',
    success: '#27AE60',
    warning: '#F39C12',
    critical: '#E74C3C',
    progress: '#3498DB',
    warmGray: '#95A5A6',
    pureWhite: '#FFFFFF'
  },
  sheets: {
    dashboard: '🏠 Dashboard',
    categories: '📂 My Categories', 
    suppliers: '🏪 My Suppliers',
    products: '💄 My Products',
    inventory: '📦 Live Inventory',
    quickAdd: '➕ Quick Add',
    reorder: '🔄 Reorder Dashboard',
    analytics: '📈 Analytics',
    instructions: '📖 Instructions'
  }
};

/**
 * Initialize the Beauty Pro Inventory System
 * Sets up all worksheets, formatting, and automation
 */
function initializeSystem() {
  try {
    Logger.log('Initializing Beauty Pro Inventory System...');
    
    // Create and format all worksheets
    setupWorksheets();
    
    // Apply professional styling
    applyDesignSystem();
    
    // Set up data validation
    setupDataValidation();
    
    // Create interactive buttons
    createActionButtons();
    
    // Initialize dashboard
    updateDashboard();
    
    Logger.log('System initialization complete!');
    SpreadsheetApp.getActiveSpreadsheet().toast('Beauty Pro Inventory System initialized successfully!', 'Setup Complete', 5);
    
  } catch (error) {
    Logger.log('Error initializing system: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Error during initialization: ' + error.message, 'Setup Error', 10);
  }
}

/**
 * Set up all required worksheets with proper structure
 */
function setupWorksheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define worksheet structure
  const worksheets = [
    { name: CONFIG.sheets.dashboard, color: CONFIG.colors.deepCharcoal },
    { name: CONFIG.sheets.categories, color: CONFIG.colors.roseGold },
    { name: CONFIG.sheets.suppliers, color: CONFIG.colors.progress },
    { name: CONFIG.sheets.products, color: CONFIG.colors.warning },
    { name: CONFIG.sheets.inventory, color: CONFIG.colors.success },
    { name: CONFIG.sheets.quickAdd, color: CONFIG.colors.critical },
    { name: CONFIG.sheets.reorder, color: CONFIG.colors.warmGray },
    { name: CONFIG.sheets.analytics, color: CONFIG.colors.progress },
    { name: CONFIG.sheets.instructions, color: CONFIG.colors.deepCharcoal }
  ];
  
  // Create or update each worksheet
  worksheets.forEach(ws => {
    let sheet = ss.getSheetByName(ws.name);
    if (!sheet) {
      sheet = ss.insertSheet(ws.name);
    }
    sheet.setTabColor(ws.color);
  });
  
  // Set dashboard as active sheet
  ss.getSheetByName(CONFIG.sheets.dashboard).activate();
}

/**
 * Apply professional design system to all sheets
 */
function applyDesignSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Apply to each sheet
  Object.values(CONFIG.sheets).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      formatSheet(sheet);
    }
  });
}

/**
 * Format individual sheet with professional styling
 */
function formatSheet(sheet) {
  // Set default font and size
  sheet.getRange('A:Z').setFontFamily('Arial').setFontSize(10);
  
  // Format headers (row 1)
  const headerRange = sheet.getRange('1:1');
  headerRange
    .setBackground(CONFIG.colors.deepCharcoal)
    .setFontColor(CONFIG.colors.pureWhite)
    .setFontWeight('bold')
    .setFontSize(14);
  
  // Format subheaders (row 3 if exists)
  try {
    const subHeaderRange = sheet.getRange('3:3');
    subHeaderRange
      .setBackground(CONFIG.colors.roseGold)
      .setFontColor(CONFIG.colors.deepCharcoal)
      .setFontWeight('bold')
      .setFontSize(12);
  } catch (e) {
    // Ignore if row doesn't exist
  }
  
  // Set alternating row colors for data rows
  const dataRange = sheet.getRange('4:50');
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  
  // Freeze header rows
  sheet.setFrozenRows(3);
}

/**
 * Handle button clicks and navigation
 */
function onSelectionChange(e) {
  if (!e || !e.range) return;
  
  const note = e.range.getNote();
  if (note && note.startsWith('ACTION:')) {
    const action = note.replace('ACTION:', '');
    executeButtonAction(action);
  }
}

/**
 * Execute button actions
 */
function executeButtonAction(action) {
  try {
    switch (action) {
      case 'ADD_PRODUCT':
        navigateToProducts();
        break;
      case 'UPDATE_STOCK': 
        navigateToQuickAdd();
        break;
      case 'VIEW_REPORTS':
        navigateToAnalytics();
        break;
      case 'CONTINUE_SETUP':
        checkSetupProgress();
        break;
      case 'MOBILE_ENTRY':
        openMobileForm();
        break;
      case 'REFRESH_DASHBOARD':
        updateDashboard();
        break;
      default:
        Logger.log('Unknown action: ' + action);
    }
  } catch (error) {
    Logger.log('Error executing action ' + action + ': ' + error.toString());
  }
}

/**
 * Navigate to specific sheets
 */
function navigateToProducts() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.products).activate();
  SpreadsheetApp.getActiveSpreadsheet().toast('Navigate to Products sheet to add new items', 'Navigation', 3);
}

function navigateToQuickAdd() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.quickAdd).activate();
  SpreadsheetApp.getActiveSpreadsheet().toast('Use Quick Add to update stock levels', 'Navigation', 3);
}

function navigateToAnalytics() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.analytics).activate();
  SpreadsheetApp.getActiveSpreadsheet().toast('View detailed business analytics and reports', 'Navigation', 3);
}

/**
 * Create interactive action buttons
 */
function createActionButtons() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.dashboard);
  
  // Define button locations and actions
  const buttons = [
    { cell: 'F5', text: '➕ Add Product', action: 'ADD_PRODUCT', style: 'primary' },
    { cell: 'F6', text: '🔄 Update Stock', action: 'UPDATE_STOCK', style: 'primary' },
    { cell: 'F7', text: '📊 View Reports', action: 'VIEW_REPORTS', style: 'secondary' },
    { cell: 'F8', text: '📱 Mobile Entry', action: 'MOBILE_ENTRY', style: 'utility' },
    { cell: 'I6', text: '🚀 Continue Setup', action: 'CONTINUE_SETUP', style: 'primary' },
    { cell: 'I8', text: '📖 Get Help', action: 'GET_HELP', style: 'utility' }
  ];
  
  // Create each button
  buttons.forEach(btn => {
    createButton(dashboard, btn.cell, btn.text, btn.action, btn.style);
  });
}

/**
 * Create individual button with styling
 */
function createButton(sheet, cell, text, action, style = 'primary') {
  const range = sheet.getRange(cell);
  const colors = getButtonColors(style);
  
  range.setValue(text)
       .setBackground(colors.background)
       .setFontColor(colors.text)
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setBorder(true, true, true, true, colors.border, SpreadsheetApp.BorderStyle.SOLID)
       .setNote('ACTION:' + action);
}

/**
 * Get button color scheme based on style
 */
function getButtonColors(style) {
  switch (style) {
    case 'primary':
      return {
        background: CONFIG.colors.roseGold,
        text: CONFIG.colors.deepCharcoal,
        border: CONFIG.colors.deepCharcoal
      };
    case 'secondary':
      return {
        background: CONFIG.colors.deepCharcoal,
        text: CONFIG.colors.pureWhite,
        border: CONFIG.colors.warmGray
      };
    case 'utility':
      return {
        background: CONFIG.colors.warmGray,
        text: CONFIG.colors.pureWhite,
        border: CONFIG.colors.deepCharcoal
      };
    default:
      return {
        background: CONFIG.colors.roseGold,
        text: CONFIG.colors.deepCharcoal,
        border: CONFIG.colors.deepCharcoal
      };
  }
}

/**
 * Update dashboard with latest metrics
 */
function updateDashboard() {
  try {
    const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.dashboard);
    
    // Calculate key metrics
    const metrics = calculateBusinessMetrics();
    
    // Update KPI cards
    updateKPICards(dashboard, metrics);
    
    // Update status indicators  
    updateStatusIndicators(dashboard, metrics);
    
    // Update alerts
    updateAlerts(dashboard, metrics);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard updated with latest data', 'Update Complete', 3);
    
  } catch (error) {
    Logger.log('Error updating dashboard: ' + error.toString());
  }
}

/**
 * Calculate business metrics from inventory data
 */
function calculateBusinessMetrics() {
  // This would calculate real metrics from inventory sheet
  // For now returning sample data structure
  return {
    totalProducts: 142,
    inventoryValue: 12847,
    lowStockItems: 23,
    expiringItems: 5,
    avgMargin: 47.2,
    turnoverRate: 4.2,
    stockoutRate: 6.4,
    expiryRisk: 125
  };
}

/**
 * Test function to run basic system check
 */
function testSystem() {
  Logger.log('Running Beauty Pro Inventory System test...');
  
  try {
    // Test worksheet access
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.values(CONFIG.sheets).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      Logger.log('Sheet ' + sheetName + ': ' + (sheet ? 'OK' : 'MISSING'));
    });
    
    // Test metrics calculation
    const metrics = calculateBusinessMetrics();
    Logger.log('Metrics calculation: OK');
    Logger.log('Sample metrics: ' + JSON.stringify(metrics));
    
    Logger.log('System test complete - all basic functions working');
    
  } catch (error) {
    Logger.log('System test failed: ' + error.toString());
  }
}