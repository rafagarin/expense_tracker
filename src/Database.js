/**
 * Database operations for the expense tracker
 * Handles all interactions with the Google Sheets database
 */

class Database {
  constructor() {
    this.sheet = null;
    this.initializeSheet();
  }

  /**
   * Initialize the spreadsheet and get the Movements sheet
   */
  initializeSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = ss.getSheetByName(SPREADSHEET_TAB_NAME);
    
    if (!this.sheet) {
      throw new Error(`Sheet named "${SPREADSHEET_TAB_NAME}" not found.`);
    }
  }

  /**
   * Get all existing Gmail IDs for idempotency checking
   * @returns {Set} Set of existing Gmail IDs
   */
  getExistingGmailIds() {
    const existingGmailIds = new Set();
    
    if (this.sheet.getLastRow() > 1) {
      const gmailIdRange = this.sheet.getRange(2, COLUMNS.GMAIL_ID + 1, this.sheet.getLastRow() - 1, 1);
      const gmailIdValues = gmailIdRange.getValues();
      
      gmailIdValues.forEach(row => {
        if (row[0]) {
          existingGmailIds.add(row[0]);
        }
      });
    }
    
    Logger.log(`Found ${existingGmailIds.size} existing movement(s) in the sheet.`);
    return existingGmailIds;
  }

  /**
   * Get existing accounting system IDs for idempotency checking
   * @returns {Set} Set of existing accounting system IDs
   */
  getExistingAccountingSystemIds() {
    const existingAccountingSystemIds = new Set();
    
    if (this.sheet.getLastRow() > 1) {
      const accountingSystemIdRange = this.sheet.getRange(2, COLUMNS.ACCOUNTING_SYSTEM_ID + 1, this.sheet.getLastRow() - 1, 1);
      const accountingSystemIdValues = accountingSystemIdRange.getValues();
      
      accountingSystemIdValues.forEach(row => {
        if (row[0]) {
          const idStr = row[0].toString();
          existingAccountingSystemIds.add(idStr);
          Logger.log(`Found existing accounting system ID: ${idStr} (type: ${typeof idStr})`);
        }
      });
    }
    
    Logger.log(`Found ${existingAccountingSystemIds.size} existing accounting system movement(s) in the sheet.`);
    return existingAccountingSystemIds;
  }

  /**
   * Add a new movement to the database
   * @param {Array} movementRow - Array representing the movement data
   */
  addMovement(movementRow) {
    // Validate that the ID is unique before adding
    const movementId = movementRow[COLUMNS.ID];
    if (this.idExists(movementId)) {
      throw new Error(`Movement with ID ${movementId} already exists in the database. This would create a duplicate ID.`);
    }
    
    this.sheet.appendRow(movementRow);
  }

  /**
   * Add multiple movements to the database in batch
   * @param {Array} movements - Array of movement objects with {ts, row, gmailId}
   */
  addMovementsBatch(movements) {
    // Sort by timestamp before adding
    const sortedMovements = movements.sort((a, b) => {
      const ta = a.ts ? Date.parse(a.ts) : 0;
      const tb = b.ts ? Date.parse(b.ts) : 0;
      return ta - tb;
    });

    sortedMovements.forEach(item => {
      this.addMovement(item.row);
      Logger.log(`Added movement from email ID: ${item.gmailId} for ${item.row[COLUMNS.CURRENCY]} ${item.row[COLUMNS.AMOUNT]} at ${item.row[COLUMNS.TIMESTAMP]} — ${item.row[COLUMNS.SOURCE_DESCRIPTION]}`);
    });
  }

  /**
   * Get the next available ID for a new movement
   * @returns {number} Next available ID
   */
  getNextId() {
    if (this.sheet.getLastRow() <= 1) {
      // If no data rows exist, start with ID 1
      return 1;
    }
    
    // Get all existing IDs from the database
    const idRange = this.sheet.getRange(2, COLUMNS.ID + 1, this.sheet.getLastRow() - 1, 1);
    const idValues = idRange.getValues();
    
    // Find the maximum ID value
    let maxId = 0;
    idValues.forEach(row => {
      const id = row[0];
      if (id && !isNaN(id) && id > maxId) {
        maxId = id;
      }
    });
    
    // Return the next available ID
    return maxId + 1;
  }

  /**
   * Check if an ID already exists in the database
   * @param {number} id - The ID to check
   * @returns {boolean} True if the ID exists, false otherwise
   */
  idExists(id) {
    if (this.sheet.getLastRow() <= 1) {
      return false;
    }
    
    const idRange = this.sheet.getRange(2, COLUMNS.ID + 1, this.sheet.getLastRow() - 1, 1);
    const idValues = idRange.getValues();
    
    return idValues.some(row => row[0] === id);
  }

  /**
   * Get all movements from the database
   * @returns {Array} Array of movement data
   */
  getAllMovements() {
    if (this.sheet.getLastRow() <= 1) {
      return [];
    }
    
    const dataRange = this.sheet.getRange(2, 1, this.sheet.getLastRow() - 1, this.sheet.getLastColumn());
    return dataRange.getValues();
  }

  /**
   * Get movements by Gmail ID
   * @param {string} gmailId - Gmail ID to search for
   * @returns {Array} Array of matching movements
   */
  getMovementsByGmailId(gmailId) {
    const allMovements = this.getAllMovements();
    return allMovements.filter(movement => movement[COLUMNS.GMAIL_ID] === gmailId);
  }

  /**
   * Get movements by accounting system ID
   * @param {string} accountingSystemId - Accounting system ID to search for
   * @returns {Array} Array of matching movements
   */
  getMovementsByAccountingSystemId(accountingSystemId) {
    const allMovements = this.getAllMovements();
    return allMovements.filter(movement => movement[COLUMNS.ACCOUNTING_SYSTEM_ID] === accountingSystemId);
  }

  /**
   * Get movements that have user_description but no category
   * @returns {Array} Array of movements needing category analysis
   */
  getMovementsNeedingCategoryAnalysis() {
    const allMovements = this.getAllMovements();
    return allMovements.filter(movement => {
      const hasUserDescription = movement[COLUMNS.USER_DESCRIPTION] && movement[COLUMNS.USER_DESCRIPTION].trim() !== '';
      const hasNoCategory = !movement[COLUMNS.CATEGORY] || movement[COLUMNS.CATEGORY].trim() === '';
      return hasUserDescription && hasNoCategory;
    });
  }

  /**
   * Update the category for a specific movement by ID
   * @param {number} movementId - The ID of the movement to update
   * @param {string} category - The new category value
   */
  updateMovementCategory(movementId, category) {
    // Find the row that contains the movement with the given ID
    const allMovements = this.getAllMovements();
    const movementRowIndex = allMovements.findIndex(movement => movement[COLUMNS.ID] === movementId);
    
    if (movementRowIndex === -1) {
      Logger.log(`Movement with ID ${movementId} not found in database`);
      return;
    }
    
    // Convert to 1-based row index (add 2 because getAllMovements() starts from row 2, and arrays are 0-based)
    const sheetRowIndex = movementRowIndex + 2;
    const categoryColumn = COLUMNS.CATEGORY + 1; // Convert to 1-based column index
    
    this.sheet.getRange(sheetRowIndex, categoryColumn).setValue(category);
    Logger.log(`Updated category for movement ID ${movementId} (row ${sheetRowIndex}) to: ${category}`);
  }

  /**
   * Update movement with analysis results
   * @param {number} movementId - The ID of the movement to update
   * @param {Object} analysisResult - Analysis result with category and split information
   */
  updateMovementWithAnalysis(movementId, analysisResult) {
    // Find the row that contains the movement with the given ID
    const allMovements = this.getAllMovements();
    const movementRowIndex = allMovements.findIndex(movement => movement[COLUMNS.ID] === movementId);
    
    if (movementRowIndex === -1) {
      Logger.log(`Movement with ID ${movementId} not found in database`);
      return;
    }
    
    // Convert to 1-based row index (add 2 because getAllMovements() starts from row 2, and arrays are 0-based)
    const sheetRowIndex = movementRowIndex + 2;
    
    // Update category
    this.sheet.getRange(sheetRowIndex, COLUMNS.CATEGORY + 1).setValue(analysisResult.category);
    
    // Update user description with clean description
    this.sheet.getRange(sheetRowIndex, COLUMNS.USER_DESCRIPTION + 1).setValue(analysisResult.clean_description);
    
    // Update comment with split instructions if available
    if (analysisResult.split_instructions) {
      this.sheet.getRange(sheetRowIndex, COLUMNS.COMMENT + 1).setValue(analysisResult.split_instructions);
    }
    
    Logger.log(`Updated movement ID ${movementId} with analysis results: category=${analysisResult.category}, needs_split=${analysisResult.needs_split}`);
  }

  /**
   * Split a movement by modifying the original row and creating a new debit row
   * @param {number} originalMovementId - The ID of the original movement to split
   * @param {Object} splitInfo - Split information from AI analysis
   * @returns {number} The ID of the new debit movement
   */
  splitMovement(originalMovementId, splitInfo) {
    // Get the original movement
    const allMovements = this.getAllMovements();
    const originalMovementIndex = allMovements.findIndex(movement => movement[COLUMNS.ID] === originalMovementId);
    
    if (originalMovementIndex === -1) {
      Logger.log(`Original movement with ID ${originalMovementId} not found in database`);
      return null;
    }
    
    const originalMovement = allMovements[originalMovementIndex];
    const nextId = this.getNextId();
    
    // Calculate the remaining amount (original - split amount)
    const remainingAmount = originalMovement[COLUMNS.AMOUNT] - splitInfo.split_amount;
    
    // Get currency conversion service for proportional splitting
    const currencyConversionService = new CurrencyConversionService();
    
    // Get original currency values
    const originalCurrencyValues = {
      clpValue: originalMovement[COLUMNS.CLP_VALUE],
      usdValue: originalMovement[COLUMNS.USD_VALUE],
      gbpValue: originalMovement[COLUMNS.GBP_VALUE]
    };
    
    // Calculate split ratio
    const splitRatio = splitInfo.split_amount / originalMovement[COLUMNS.AMOUNT];
    
    // Split currency values proportionally for personal portion
    const personalCurrencyValues = currencyConversionService.splitCurrencyValues(
      originalMovement[COLUMNS.AMOUNT],
      splitInfo.split_amount,
      originalCurrencyValues
    );
    
    // Split currency values proportionally for shared portion
    const sharedCurrencyValues = currencyConversionService.splitCurrencyValues(
      originalMovement[COLUMNS.AMOUNT],
      remainingAmount,
      originalCurrencyValues
    );
    
    // 1. Modify the original row in place to be the personal portion
    const sheetRowIndex = originalMovementIndex + 2;
    this.sheet.getRange(sheetRowIndex, COLUMNS.AMOUNT + 1).setValue(splitInfo.split_amount);
    this.sheet.getRange(sheetRowIndex, COLUMNS.CATEGORY + 1).setValue(splitInfo.split_category);
    this.sheet.getRange(sheetRowIndex, COLUMNS.COMMENT + 1).setValue(''); // Clear comment for expense line
    
    // Update currency values for the personal portion
    this.sheet.getRange(sheetRowIndex, COLUMNS.CLP_VALUE + 1).setValue(personalCurrencyValues.clpValue);
    this.sheet.getRange(sheetRowIndex, COLUMNS.USD_VALUE + 1).setValue(personalCurrencyValues.usdValue);
    this.sheet.getRange(sheetRowIndex, COLUMNS.GBP_VALUE + 1).setValue(personalCurrencyValues.gbpValue);
    
    // 2. Create the shared portion as a new debit movement
    const sharedMovement = [...originalMovement];
    sharedMovement[COLUMNS.ID] = nextId;
    sharedMovement[COLUMNS.AMOUNT] = remainingAmount;
    sharedMovement[COLUMNS.CATEGORY] = originalMovement[COLUMNS.CATEGORY];
    sharedMovement[COLUMNS.USER_DESCRIPTION] = originalMovement[COLUMNS.USER_DESCRIPTION];
    sharedMovement[COLUMNS.DIRECTION] = DIRECTIONS.NEUTRAL;
    sharedMovement[COLUMNS.TYPE] = MOVEMENT_TYPES.DEBIT;
    sharedMovement[COLUMNS.STATUS] = STATUS.PENDING_DIRECT_SETTLEMENT;
    sharedMovement[COLUMNS.SOURCE] = SOURCES.GMAIL;
    sharedMovement[COLUMNS.COMMENT] = `Split from #${originalMovementId}`; // Reference the expense line
    
    // Set currency values for the shared portion
    sharedMovement[COLUMNS.CLP_VALUE] = sharedCurrencyValues.clpValue;
    sharedMovement[COLUMNS.USD_VALUE] = sharedCurrencyValues.usdValue;
    sharedMovement[COLUMNS.GBP_VALUE] = sharedCurrencyValues.gbpValue;
    
    // Add the new debit movement to the database
    this.addMovement(sharedMovement);
    
    Logger.log(`Split movement ID ${originalMovementId}: modified original to personal portion (${splitInfo.split_amount}), created debit movement ${nextId} for shared portion (${remainingAmount})`);
    
    return nextId;
  }

  /**
   * Get movements that are awaiting Splitwise upload
   * @returns {Array} Array of movements awaiting Splitwise upload
   */
  getMovementsPendingSplitwiseSettlement() {
    const allMovements = this.getAllMovements();
    return allMovements.filter(movement => 
      movement[COLUMNS.STATUS] === STATUS.PENDING_SPLITWISE_SETTLEMENT
    );
  }

  /**
   * Get debit movements that are pending settlement
   * @returns {Array} Array of debit movements pending settlement
   */
  getMovementsPendingDirectSettlement() {
    const allMovements = this.getAllMovements();
    return allMovements.filter(movement => 
      movement[COLUMNS.TYPE] === MOVEMENT_TYPES.DEBIT && 
      movement[COLUMNS.STATUS] === STATUS.PENDING_DIRECT_SETTLEMENT
    );
  }

  /**
   * Update the settled_movement_id for a movement
   * @param {number} movementId - The ID of the movement to update
   * @param {number} settledMovementId - The ID of the movement that settles this one
   */
  updateMovementSettledId(movementId, settledMovementId) {
    // Find the row that contains the movement with the given ID
    const allMovements = this.getAllMovements();
    const movementRowIndex = allMovements.findIndex(movement => movement[COLUMNS.ID] === movementId);
    
    if (movementRowIndex === -1) {
      Logger.log(`Movement with ID ${movementId} not found in database`);
      return;
    }
    
    // Convert to 1-based row index (add 2 because getAllMovements() starts from row 2, and arrays are 0-based)
    const sheetRowIndex = movementRowIndex + 2;
    
    // Update settled_movement_id
    this.sheet.getRange(sheetRowIndex, COLUMNS.SETTLED_MOVEMENT_ID + 1).setValue(settledMovementId);
    
    Logger.log(`Updated movement ID ${movementId} with settled_movement_id: ${settledMovementId}`);
  }

  /**
   * Update the status for a movement
   * @param {number} movementId - The ID of the movement to update
   * @param {string} status - The new status value
   */
  updateMovementStatus(movementId, status) {
    // Find the row that contains the movement with the given ID
    const allMovements = this.getAllMovements();
    const movementRowIndex = allMovements.findIndex(movement => movement[COLUMNS.ID] === movementId);
    
    if (movementRowIndex === -1) {
      Logger.log(`Movement with ID ${movementId} not found in database`);
      return;
    }
    
    // Convert to 1-based row index (add 2 because getAllMovements() starts from row 2, and arrays are 0-based)
    const sheetRowIndex = movementRowIndex + 2;
    
    // Update status
    this.sheet.getRange(sheetRowIndex, COLUMNS.STATUS + 1).setValue(status);
    
    Logger.log(`Updated movement ID ${movementId} status to: ${status}`);
  }

  /**
   * Update movement with Splitwise information after pushing to Splitwise
   * @param {number} movementId - The ID of the movement to update
   * @param {string} splitwiseId - The Splitwise expense ID
   */
  updateMovementWithSplitwiseInfo(movementId, splitwiseId) {
    // Find the row that contains the movement with the given ID
    const allMovements = this.getAllMovements();
    const movementRowIndex = allMovements.findIndex(movement => movement[COLUMNS.ID] === movementId);
    
    if (movementRowIndex === -1) {
      Logger.log(`Movement with ID ${movementId} not found in database`);
      return;
    }
    
    // Convert to 1-based row index (add 2 because getAllMovements() starts from row 2, and arrays are 0-based)
    const sheetRowIndex = movementRowIndex + 2;
    
    // Update Splitwise information
    this.sheet.getRange(sheetRowIndex, COLUMNS.ACCOUNTING_SYSTEM_ID + 1).setValue(splitwiseId);
    this.sheet.getRange(sheetRowIndex, COLUMNS.STATUS + 1).setValue(STATUS.IN_SPLITWISE);
    
    Logger.log(`Updated movement ID ${movementId} with Splitwise ID ${splitwiseId} and marked as in splitwise`);
  }


}
