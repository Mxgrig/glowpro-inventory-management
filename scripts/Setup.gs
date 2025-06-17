/**
 * Beauty Pro Inventory System - Setup & Progress Tracking
 * Guided setup flow with progress monitoring
 */

/**
 * Check overall setup progress and guide user through next steps
 */
function checkSetupProgress() {
  try {
    const progress = {
      categories: checkCategoriesComplete(),
      suppliers: checkSuppliersComplete(),
      products: checkProductsComplete(), 
      inventory: checkInventoryComplete()
    };
    
    const overallProgress = calculateOverallProgress(progress);
    
    // Update progress display
    updateProgressDisplay(overallProgress, progress);
    
    // Show next action guidance
    showNextActionGuidance(progress);
    
    return {
      progress: overallProgress,
      steps: progress,
      nextAction: getNextAction(progress)
    };
    
  } catch (error) {
    Logger.log('Error checking setup progress: ' + error.toString());
    return { progress: 0, steps: {}, nextAction: 'error' };
  }
}

/**
 * Check if categories setup is complete
 */
function checkCategoriesComplete() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📂 My Categories');
    if (!sheet) return false;
    
    const data = sheet.getRange('A4:E20').getValues();
    const validCategories = data.filter(row => 
      row[0] && row[0] !== '' && // Category name exists
      row[1] && row[1] !== '' && // Description exists
      row[2] && row[2] !== ''    // Target margin exists
    );
    
    // Need at least 3 valid categories
    return validCategories.length >= 3;
    
  } catch (error) {
    Logger.log('Error checking categories: ' + error.toString());
    return false;
  }
}

/**
 * Check if suppliers setup is complete
 */
function checkSuppliersComplete() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🏪 My Suppliers');
    if (!sheet) return false;
    
    const data = sheet.getRange('A4:J20').getValues();
    const validSuppliers = data.filter(row => 
      row[0] && row[0] !== '' && // Supplier name
      row[1] && row[1] !== '' && // Contact person
      row[2] && row[2] !== '' && // Email
      row[4] && row[4] !== ''    // Lead time
    );
    
    // Need at least 1 complete supplier
    return validSuppliers.length >= 1;
    
  } catch (error) {
    Logger.log('Error checking suppliers: ' + error.toString());
    return false;
  }
}

/**
 * Check if products setup is complete
 */
function checkProductsComplete() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💄 My Products');
    if (!sheet) return false;
    
    const data = sheet.getRange('A4:N50').getValues();
    const validProducts = data.filter(row => 
      row[0] && row[0] !== '' && // Product name
      row[1] && row[1] !== '' && // Brand
      row[2] && row[2] !== '' && // Category
      row[5] && row[5] !== '' && // Cost
      row[6] && row[6] !== ''    // Retail price
    );
    
    // Need at least 5 complete products
    return validProducts.length >= 5;
    
  } catch (error) {
    Logger.log('Error checking products: ' + error.toString());
    return false;
  }
}

/**
 * Check if inventory setup is complete
 */
function checkInventoryComplete() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📦 Live Inventory');
    if (!sheet) return false;
    
    const data = sheet.getRange('A4:O50').getValues();
    const validInventory = data.filter(row => 
      row[0] && row[0] !== '' && // Product name
      row[1] !== '' && row[1] !== null && // Current stock (can be 0)
      row[2] && row[2] !== '' && // Min stock
      row[3] && row[3] !== ''    // Max stock
    );
    
    // Need inventory data for products
    return validInventory.length >= 3;
    
  } catch (error) {
    Logger.log('Error checking inventory: ' + error.toString());
    return false;
  }
}

/**
 * Calculate overall progress percentage
 */
function calculateOverallProgress(progress) {
  const steps = Object.values(progress);
  const completedSteps = steps.filter(step => step === true).length;
  return Math.round((completedSteps / steps.length) * 100);
}

/**
 * Update progress display on dashboard
 */
function updateProgressDisplay(overallProgress, stepProgress) {
  try {
    const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🏠 Dashboard');
    if (!dashboard) return;
    
    // Update progress percentage
    dashboard.getRange('I4').setValue(`Setup Complete: ${overallProgress}%`);
    
    // Update remaining steps
    const remainingSteps = Object.values(stepProgress).filter(step => !step).length;
    dashboard.getRange('I5').setValue(`⚠️ ${remainingSteps} Steps Remaining`);
    
    // Update progress bar color based on completion
    const progressCell = dashboard.getRange('I4');
    if (overallProgress === 100) {
      progressCell.setBackground('#27AE60').setFontColor('#FFFFFF');
    } else if (overallProgress >= 75) {
      progressCell.setBackground('#F39C12').setFontColor('#FFFFFF');
    } else {
      progressCell.setBackground('#E74C3C').setFontColor('#FFFFFF');
    }
    
  } catch (error) {
    Logger.log('Error updating progress display: ' + error.toString());
  }
}

/**
 * Show guidance for next action needed
 */
function showNextActionGuidance(progress) {
  try {
    const nextAction = getNextAction(progress);
    let message = '';
    let title = '';
    
    switch (nextAction) {
      case 'categories':
        title = 'Next: Setup Categories';
        message = 'Go to 📂 My Categories sheet and replace the example categories with your business categories. You need at least 3 categories to continue.';
        break;
      case 'suppliers':
        title = 'Next: Add Suppliers';
        message = 'Go to 🏪 My Suppliers sheet and add your vendor information. You need at least 1 complete supplier to continue.';
        break;
      case 'products':
        title = 'Next: Add Products';
        message = 'Go to 💄 My Products sheet and add your product catalog. You need at least 5 products with complete information.';
        break;
      case 'inventory':
        title = 'Next: Set Inventory Levels';
        message = 'Go to 📦 Live Inventory sheet and set current stock levels for your products.';
        break;
      case 'complete':
        title = '🎉 Setup Complete!';
        message = 'Congratulations! Your Beauty Pro Inventory System is fully configured and ready to use. Check out the Analytics sheet for business insights.';
        break;
      default:
        title = 'Setup in Progress';
        message = 'Continue working through the setup steps. Check the Instructions sheet if you need help.';
    }
    
    // Show guidance toast
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 8);
    
  } catch (error) {
    Logger.log('Error showing guidance: ' + error.toString());
  }
}

/**
 * Determine next action based on progress
 */
function getNextAction(progress) {
  if (!progress.categories) return 'categories';
  if (!progress.suppliers) return 'suppliers';
  if (!progress.products) return 'products';
  if (!progress.inventory) return 'inventory';
  return 'complete';
}

/**
 * Navigate to next incomplete step
 */
function navigateToNextStep() {
  try {
    const progress = checkSetupProgress();
    const nextAction = progress.nextAction;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    switch (nextAction) {
      case 'categories':
        ss.getSheetByName('📂 My Categories').activate();
        ss.getRange('A4').activate();
        break;
      case 'suppliers':
        ss.getSheetByName('🏪 My Suppliers').activate();
        ss.getRange('A4').activate();
        break;
      case 'products':
        ss.getSheetByName('💄 My Products').activate();
        ss.getRange('A4').activate();
        break;
      case 'inventory':
        ss.getSheetByName('📦 Live Inventory').activate();
        ss.getRange('B4').activate();
        break;
      case 'complete':
        ss.getSheetByName('📈 Analytics').activate();
        break;
      default:
        ss.getSheetByName('📖 Instructions').activate();
    }
    
  } catch (error) {
    Logger.log('Error navigating to next step: ' + error.toString());
  }
}

/**
 * Clear demo data and prepare for user setup
 */
function clearDemoData() {
  try {
    const response = SpreadsheetApp.getUi().alert(
      'Clear Demo Data',
      'This will remove all example data and prepare the system for your business. This cannot be undone. Continue?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response === SpreadsheetApp.getUi().Button.YES) {
      // Clear categories (keep headers)
      const categoriesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📂 My Categories');
      if (categoriesSheet) {
        categoriesSheet.getRange('A4:H20').clearContent();
      }
      
      // Clear suppliers (keep headers)
      const suppliersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🏪 My Suppliers');
      if (suppliersSheet) {
        suppliersSheet.getRange('A4:J20').clearContent();
      }
      
      // Clear products (keep headers)
      const productsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('💄 My Products');
      if (productsSheet) {
        productsSheet.getRange('A4:N50').clearContent();
      }
      
      // Clear inventory (keep headers)
      const inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('📦 Live Inventory');
      if (inventorySheet) {
        inventorySheet.getRange('A4:O50').clearContent();
      }
      
      // Reset progress
      checkSetupProgress();
      
      SpreadsheetApp.getActiveSpreadsheet().toast('Demo data cleared. Ready for your business setup!', 'Data Cleared', 5);
      
      // Navigate to categories to start setup
      navigateToNextStep();
    }
    
  } catch (error) {
    Logger.log('Error clearing demo data: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Error clearing data: ' + error.message, 'Error', 5);
  }
}

/**
 * Show setup completion celebration
 */
function showSetupComplete() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Show success message
    ss.toast(
      '🎉 Congratulations! Your Beauty Pro Inventory System is fully configured and ready to manage your business. Your dashboard now shows real data from your inventory!',
      'Setup Complete!',
      10
    );
    
    // Navigate to dashboard
    ss.getSheetByName('🏠 Dashboard').activate();
    
    // Update dashboard with fresh data
    updateDashboard();
    
  } catch (error) {
    Logger.log('Error showing setup complete: ' + error.toString());
  }
}