/**
 * Beauty Pro Inventory System - Data Validation Module
 * Smart dropdowns and data validation rules
 */

/**
 * Set up all data validation rules for the system
 */
function setupDataValidation() {
  try {
    Logger.log('Setting up data validation...');
    
    // Set up category dropdowns
    setupCategoryDropdowns();
    
    // Set up supplier dropdowns
    setupSupplierDropdowns();
    
    // Set up status validations
    setupStatusValidations();
    
    // Set up numeric validations
    setupNumericValidations();
    
    Logger.log('Data validation setup complete');
    SpreadsheetApp.getActiveSpreadsheet().toast('Data validation rules applied successfully', 'Validation Setup', 3);
    
  } catch (error) {
    Logger.log('Error setting up data validation: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Error setting up validation: ' + error.message, 'Validation Error', 5);
  }
}

/**
 * Set up category dropdown validation
 */
function setupCategoryDropdowns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get categories from the Categories sheet
    const categoriesSheet = ss.getSheetByName('📂 My Categories');
    if (!categoriesSheet) return;
    
    // Create category dropdown for Products sheet
    const productsSheet = ss.getSheetByName('💄 My Products');
    if (productsSheet) {
      const categoryRange = productsSheet.getRange('C4:C100'); // Category column
      const categorySource = categoriesSheet.getRange('A4:A20');
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(categorySource, true)
        .setAllowInvalid(false)
        .setHelpText('Select a category from your configured categories')
        .build();
        
      categoryRange.setDataValidation(rule);
    }
    
    Logger.log('Category dropdowns configured');
    
  } catch (error) {
    Logger.log('Error setting up category dropdowns: ' + error.toString());
  }
}

/**
 * Set up supplier dropdown validation
 */
function setupSupplierDropdowns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get suppliers from the Suppliers sheet
    const suppliersSheet = ss.getSheetByName('🏪 My Suppliers');
    if (!suppliersSheet) return;
    
    // Create supplier dropdown for Products sheet
    const productsSheet = ss.getSheetByName('💄 My Products');
    if (productsSheet) {
      const supplierRange = productsSheet.getRange('E4:E100'); // Supplier column
      const supplierSource = suppliersSheet.getRange('A4:A20');
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(supplierSource, true)
        .setAllowInvalid(false)
        .setHelpText('Select a supplier from your configured suppliers')
        .build();
        
      supplierRange.setDataValidation(rule);
    }
    
    // Also set up for Inventory sheet
    const inventorySheet = ss.getSheetByName('📦 Live Inventory');
    if (inventorySheet) {
      const inventorySupplierRange = inventorySheet.getRange('N4:N100');
      const supplierSource = suppliersSheet.getRange('A4:A20');
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(supplierSource, true)
        .setAllowInvalid(false)
        .setHelpText('Select supplier for this product')
        .build();
        
      inventorySupplierRange.setDataValidation(rule);
    }
    
    Logger.log('Supplier dropdowns configured');
    
  } catch (error) {
    Logger.log('Error setting up supplier dropdowns: ' + error.toString());
  }
}

/**
 * Set up status validation rules
 */
function setupStatusValidations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Set up status dropdowns for Categories sheet
    const categoriesSheet = ss.getSheetByName('📂 My Categories');
    if (categoriesSheet) {
      const statusRange = categoriesSheet.getRange('E4:E20');
      const statusValues = ['✅ Active', '⏸️ Inactive', '🚫 Discontinued'];
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(statusValues, true)
        .setAllowInvalid(false)
        .setHelpText('Select category status')
        .build();
        
      statusRange.setDataValidation(rule);
    }
    
    // Set up inventory status validation
    const inventorySheet = ss.getSheetByName('📦 Live Inventory');
    if (inventorySheet) {
      const inventoryStatusRange = inventorySheet.getRange('F4:F100');
      const inventoryStatusValues = ['🟢 Healthy', '🟡 At Minimum', '🔴 Low Stock', '🔴 Out of Stock', '⚫ Expired'];
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(inventoryStatusValues, true)
        .setAllowInvalid(false)
        .setHelpText('Current inventory status')
        .build();
        
      inventoryStatusRange.setDataValidation(rule);
    }
    
    // Set up Quick Add transaction types
    const quickAddSheet = ss.getSheetByName('➕ Quick Add');
    if (quickAddSheet) {
      const transactionRange = quickAddSheet.getRange('G4');
      const transactionTypes = ['Stock Received', 'Sale', 'Return', 'Damaged', 'Expired', 'Transfer', 'Adjustment'];
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(transactionTypes, true)
        .setAllowInvalid(false)
        .setHelpText('Select transaction type')
        .build();
        
      transactionRange.setDataValidation(rule);
    }
    
    Logger.log('Status validations configured');
    
  } catch (error) {
    Logger.log('Error setting up status validations: ' + error.toString());
  }
}

/**
 * Set up numeric validation rules
 */
function setupNumericValidations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Set up percentage validation for margins
    const categoriesSheet = ss.getSheetByName('📂 My Categories');
    if (categoriesSheet) {
      const marginRange = categoriesSheet.getRange('C4:C20');
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireNumberBetween(0, 1)
        .setAllowInvalid(false)
        .setHelpText('Enter margin as decimal (e.g., 0.45 for 45%)')
        .build();
        
      marginRange.setDataValidation(rule);
      marginRange.setNumberFormat('0.0%');
    }
    
    // Set up positive number validation for costs and prices
    const productsSheet = ss.getSheetByName('💄 My Products');
    if (productsSheet) {
      // Cost validation
      const costRange = productsSheet.getRange('F4:F100');
      const costRule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(false)
        .setHelpText('Enter product cost (must be greater than 0)')
        .build();
      costRange.setDataValidation(costRule);
      costRange.setNumberFormat('$#,##0.00');
      
      // Retail price validation
      const priceRange = productsSheet.getRange('G4:G100');
      const priceRule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(false)
        .setHelpText('Enter retail price (must be greater than 0)')
        .build();
      priceRange.setDataValidation(priceRule);
      priceRange.setNumberFormat('$#,##0.00');
    }
    
    // Set up stock level validations
    const inventorySheet = ss.getSheetByName('📦 Live Inventory');
    if (inventorySheet) {
      // Current stock validation
      const stockRange = inventorySheet.getRange('B4:B100');
      const stockRule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0)
        .setAllowInvalid(false)
        .setHelpText('Enter current stock quantity (0 or greater)')
        .build();
      stockRange.setDataValidation(stockRule);
      
      // Min stock validation
      const minStockRange = inventorySheet.getRange('C4:C100');
      const minStockRule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(false)
        .setHelpText('Enter minimum stock level (must be greater than 0)')
        .build();
      minStockRange.setDataValidation(minStockRule);
      
      // Max stock validation
      const maxStockRange = inventorySheet.getRange('D4:D100');
      const maxStockRule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThan(0)
        .setAllowInvalid(false)
        .setHelpText('Enter maximum stock level (must be greater than 0)')
        .build();
      maxStockRange.setDataValidation(maxStockRule);
    }
    
    Logger.log('Numeric validations configured');
    
  } catch (error) {
    Logger.log('Error setting up numeric validations: ' + error.toString());
  }
}

/**
 * Set up date validation rules
 */
function setupDateValidations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Set up expiry date validation in Products sheet
    const productsSheet = ss.getSheetByName('💄 My Products');
    if (productsSheet) {
      const expiryRange = productsSheet.getRange('J4:J100');
      const today = new Date();
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireDateAfter(today)
        .setAllowInvalid(false)
        .setHelpText('Enter expiry date (must be in the future)')
        .build();
        
      expiryRange.setDataValidation(rule);
      expiryRange.setNumberFormat('mm/dd/yyyy');
    }
    
    Logger.log('Date validations configured');
    
  } catch (error) {
    Logger.log('Error setting up date validations: ' + error.toString());
  }
}

/**
 * Validate email addresses in supplier sheet
 */
function setupEmailValidation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const suppliersSheet = ss.getSheetByName('🏪 My Suppliers');
    
    if (suppliersSheet) {
      const emailRange = suppliersSheet.getRange('C4:C20');
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireTextIsEmail()
        .setAllowInvalid(false)
        .setHelpText('Enter valid email address')
        .build();
        
      emailRange.setDataValidation(rule);
    }
    
    Logger.log('Email validation configured');
    
  } catch (error) {
    Logger.log('Error setting up email validation: ' + error.toString());
  }
}

/**
 * Create smart dropdown for product selection in Quick Add
 */
function setupProductDropdown() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get products from Products sheet
    const productsSheet = ss.getSheetByName('💄 My Products');
    const quickAddSheet = ss.getSheetByName('➕ Quick Add');
    
    if (productsSheet && quickAddSheet) {
      const productRange = quickAddSheet.getRange('B4');
      const productSource = productsSheet.getRange('A4:A100');
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(productSource, true)
        .setAllowInvalid(false)
        .setHelpText('Select product to update')
        .build();
        
      productRange.setDataValidation(rule);
    }
    
    Logger.log('Product dropdown configured');
    
  } catch (error) {
    Logger.log('Error setting up product dropdown: ' + error.toString());
  }
}

/**
 * Remove all data validation (for troubleshooting)
 */
function clearAllValidation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    sheets.forEach(sheet => {
      const range = sheet.getDataRange();
      range.clearDataValidations();
    });
    
    Logger.log('All data validation cleared');
    SpreadsheetApp.getActiveSpreadsheet().toast('All data validation rules cleared', 'Validation Cleared', 3);
    
  } catch (error) {
    Logger.log('Error clearing validation: ' + error.toString());
  }
}

/**
 * Test data validation setup
 */
function testDataValidation() {
  try {
    Logger.log('Testing data validation setup...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Test category validation
    const productsSheet = ss.getSheetByName('💄 My Products');
    if (productsSheet) {
      const categoryCell = productsSheet.getRange('C4');
      const validation = categoryCell.getDataValidation();
      Logger.log('Category validation: ' + (validation ? 'OK' : 'MISSING'));
    }
    
    // Test supplier validation
    if (productsSheet) {
      const supplierCell = productsSheet.getRange('E4');
      const validation = supplierCell.getDataValidation();
      Logger.log('Supplier validation: ' + (validation ? 'OK' : 'MISSING'));
    }
    
    Logger.log('Data validation test complete');
    
  } catch (error) {
    Logger.log('Error testing data validation: ' + error.toString());
  }
}