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
   * Get all existing source IDs for idempotency checking
   * @param {string} source - The source to filter by (optional)
   * @returns {Set} Set of existing source IDs
   */
  getExistingSourceIds(source = null) {
    const existingSourceIds = new Set();
    
    if (this.sheet.getLastRow() > 1) {
      const sourceIdRange = this.sheet.getRange(2, COLUMNS.SOURCE_ID + 1, this.sheet.getLastRow() - 1, 1);
      const sourceIdValues = sourceIdRange.getValues();
      
      // If source is specified, we need to check both source_id and source columns
      if (source) {
        const allMovements = this.getAllMovements();
        allMovements.forEach((movement, index) => {
          if (movement[COLUMNS.SOURCE] === source && movement[COLUMNS.SOURCE_ID]) {
            existingSourceIds.add(movement[COLUMNS.SOURCE_ID]);
          }
        });
      } else {
        // Get all source IDs regardless of source
        sourceIdValues.forEach(row => {
          if (row[0]) {
            existingSourceIds.add(row[0]);
          }
        });
      }
    }
    
    const sourceText = source ? ` for source '${source}'` : '';
    Logger.log(`Found ${existingSourceIds.size} existing source ID(s)${sourceText} in the sheet.`);
    return existingSourceIds;
  }

  /**
   * Get all existing Gmail IDs for idempotency checking (legacy method)
   * @returns {Set} Set of existing Gmail IDs
   */
  getExistingGmailIds() {
    return this.getExistingSourceIds(SOURCES.GMAIL);
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
   * @param {Array} movements - Array of movement objects with {ts, row, sourceId}
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
      const sourceId = item.gmailId || item.monzoId || item.accountingSystemId || 'unknown';
      Logger.log(`Added movement from source ID: ${sourceId} for ${item.row[COLUMNS.CURRENCY]} ${item.row[COLUMNS.AMOUNT]} at ${item.row[COLUMNS.TIMESTAMP]} — ${item.row[COLUMNS.SOURCE_DESCRIPTION]}`);
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
    
    // If the movement is not being split, we can update the user description to the cleaned
    // version from the AI and clear the comment field, as it has been processed.
    // For split movements, this is handled within the respective split functions.
    if (!analysisResult.needs_split && analysisResult.clean_description) {
      this.sheet.getRange(sheetRowIndex, COLUMNS.USER_DESCRIPTION + 1).setValue(analysisResult.clean_description);
      this.sheet.getRange(sheetRowIndex, COLUMNS.COMMENT + 1).setValue('');
    }
    
    Logger.log(`Updated movement ID ${movementId} with analysis results: category=${analysisResult.category}`);
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
    // The category is set by updateMovementWithAnalysis, but we ensure it's correct here.
    this.sheet.getRange(sheetRowIndex, COLUMNS.CATEGORY + 1).setValue(splitInfo.split_category);
    this.sheet.getRange(sheetRowIndex, COLUMNS.USER_DESCRIPTION + 1).setValue(splitInfo.clean_description);
    this.sheet.getRange(sheetRowIndex, COLUMNS.COMMENT + 1).setValue(''); // Clear comment
    this.sheet.getRange(sheetRowIndex, COLUMNS.AI_COMMENT + 1).setValue(`Split into #${nextId}`); // Reference the debit line it was split into
    this.sheet.getRange(sheetRowIndex, COLUMNS.ORIGINAL_AMOUNT + 1).setValue(originalMovement[COLUMNS.AMOUNT]); // Set original amount
    
    // Update currency values for the personal portion
    this.sheet.getRange(sheetRowIndex, COLUMNS.CLP_VALUE + 1).setValue(personalCurrencyValues.clpValue);
    this.sheet.getRange(sheetRowIndex, COLUMNS.USD_VALUE + 1).setValue(personalCurrencyValues.usdValue);
    this.sheet.getRange(sheetRowIndex, COLUMNS.GBP_VALUE + 1).setValue(personalCurrencyValues.gbpValue);
    
    // 2. Create the shared portion as a new debit movement
    const sharedMovement = [...originalMovement];
    sharedMovement[COLUMNS.ID] = nextId;
    sharedMovement[COLUMNS.AMOUNT] = remainingAmount;
    sharedMovement[COLUMNS.CATEGORY] = splitInfo.split_category; // Keep category consistent
    sharedMovement[COLUMNS.USER_DESCRIPTION] = splitInfo.split_description; // Use description for the debit part
    sharedMovement[COLUMNS.DIRECTION] = DIRECTIONS.NEUTRAL;
    sharedMovement[COLUMNS.TYPE] = MOVEMENT_TYPES.DEBIT;
    sharedMovement[COLUMNS.STATUS] = STATUS.PENDING_DIRECT_SETTLEMENT;
    sharedMovement[COLUMNS.COMMENT] = ''; // Clear comment for debit line
    sharedMovement[COLUMNS.AI_COMMENT] = `Split from #${originalMovementId}`; // Reference the expense line in AI comment
    sharedMovement[COLUMNS.ORIGINAL_AMOUNT] = originalMovement[COLUMNS.AMOUNT]; // Set original amount
    
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
   * Splits an expense movement into two for re-categorization purposes.
   * The original movement's amount is reduced, and a new uncategorized movement is created.
   * Both resulting movements are left uncategorized for the user to detail further.
   * @param {number} originalMovementId - The ID of the original movement to split.
   * @param {Object} splitInfo - Information for the split, including split_amount and descriptions.
   * @returns {number|null} The ID of the new movement, or null on failure.
   */
  splitExpenseForRecategorization(originalMovementId, splitInfo) {
    // 1. Find the original movement
    const allMovements = this.getAllMovements();
    const originalMovementIndex = allMovements.findIndex(movement => movement[COLUMNS.ID] === originalMovementId);
    
    if (originalMovementIndex === -1) {
      Logger.log(`Original movement with ID ${originalMovementId} not found for expense split.`);
      return null;
    }
    
    const originalMovement = allMovements[originalMovementIndex];
    const nextId = this.getNextId();
    
    // 2. Validate amounts
    const originalAmount = originalMovement[COLUMNS.AMOUNT];
    const splitAmount = splitInfo.split_amount;
    if (splitAmount >= originalAmount || splitAmount <= 0) {
      Logger.log(`Invalid split amount ${splitAmount} for original amount ${originalAmount}. Aborting split.`);
      return null;
    }
    const remainingAmount = originalAmount - splitAmount;

    // 3. Handle currency value splitting
    const currencyConversionService = new CurrencyConversionService();
    const originalCurrencyValues = {
      clpValue: originalMovement[COLUMNS.CLP_VALUE],
      usdValue: originalMovement[COLUMNS.USD_VALUE],
      gbpValue: originalMovement[COLUMNS.GBP_VALUE]
    };
    
    // Proportional values for the remaining part of the original movement
    const remainingCurrencyValues = currencyConversionService.splitCurrencyValues(
      originalAmount,
      remainingAmount,
      originalCurrencyValues
    );
    
    // Proportional values for the new split-off movement
    const splitCurrencyValues = currencyConversionService.splitCurrencyValues(
      originalAmount,
      splitAmount,
      originalCurrencyValues
    );

    // 4. Modify the original movement row
    const sheetRowIndex = originalMovementIndex + 2; // 1-based index for sheet
    this.sheet.getRange(sheetRowIndex, COLUMNS.AMOUNT + 1).setValue(remainingAmount);
    this.sheet.getRange(sheetRowIndex, COLUMNS.USER_DESCRIPTION + 1).setValue(splitInfo.clean_description || originalMovement[COLUMNS.USER_DESCRIPTION]);
    this.sheet.getRange(sheetRowIndex, COLUMNS.CATEGORY + 1).setValue(null); // Un-categorize
    this.sheet.getRange(sheetRowIndex, COLUMNS.COMMENT + 1).setValue(''); // Clear comment
    this.sheet.getRange(sheetRowIndex, COLUMNS.AI_COMMENT + 1).setValue(`Split into #${nextId}`);
    this.sheet.getRange(sheetRowIndex, COLUMNS.ORIGINAL_AMOUNT + 1).setValue(originalAmount);
    
    this.sheet.getRange(sheetRowIndex, COLUMNS.CLP_VALUE + 1).setValue(remainingCurrencyValues.clpValue);
    this.sheet.getRange(sheetRowIndex, COLUMNS.USD_VALUE + 1).setValue(remainingCurrencyValues.usdValue);
    this.sheet.getRange(sheetRowIndex, COLUMNS.GBP_VALUE + 1).setValue(remainingCurrencyValues.gbpValue);

    // 5. Create the new movement row
    const newMovement = [...originalMovement];
    newMovement[COLUMNS.ID] = nextId;
    newMovement[COLUMNS.AMOUNT] = splitAmount;
    newMovement[COLUMNS.USER_DESCRIPTION] = splitInfo.split_description;
    newMovement[COLUMNS.CATEGORY] = null; // Uncategorized
    newMovement[COLUMNS.COMMENT] = ''; // Clear comment
    newMovement[COLUMNS.AI_COMMENT] = `Split from #${originalMovementId}`;
    newMovement[COLUMNS.ORIGINAL_AMOUNT] = originalAmount;
    
    newMovement[COLUMNS.CLP_VALUE] = splitCurrencyValues.clpValue;
    newMovement[COLUMNS.USD_VALUE] = splitCurrencyValues.usdValue;
    newMovement[COLUMNS.GBP_VALUE] = splitCurrencyValues.gbpValue;

    // 6. Add the new movement to the database
    this.addMovement(newMovement);
    
    Logger.log(`Split expense movement ID ${originalMovementId}: modified original to ${remainingAmount}, created new movement ${nextId} for ${splitAmount}`);
    
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

  /**
   * Get movements that have failed currency conversions (#NUM! errors)
   * @returns {Array} Array of movements with failed currency conversions
   */
  getMovementsWithFailedCurrencyConversion() {
    const allMovements = this.getAllMovements();
    const currencyConversionService = new CurrencyConversionService();
    
    return allMovements.filter(movement => {
      const clpValue = movement[COLUMNS.CLP_VALUE];
      const usdValue = movement[COLUMNS.USD_VALUE];
      const gbpValue = movement[COLUMNS.GBP_VALUE];
      
      // Check if any of the currency values represent failed conversions
      return currencyConversionService.isFailedConversion(clpValue) ||
             currencyConversionService.isFailedConversion(usdValue) ||
             currencyConversionService.isFailedConversion(gbpValue);
    });
  }

  /**
   * Fix currency conversion for a specific movement
   * @param {number} movementId - The ID of the movement to fix
   * @returns {boolean} True if the conversion was successfully fixed
   */
  fixMovementCurrencyConversion(movementId) {
    // Find the row that contains the movement with the given ID
    const allMovements = this.getAllMovements();
    const movementRowIndex = allMovements.findIndex(movement => movement[COLUMNS.ID] === movementId);
    
    if (movementRowIndex === -1) {
      Logger.log(`Movement with ID ${movementId} not found in database`);
      return false;
    }
    
    const movement = allMovements[movementRowIndex];
    const amount = movement[COLUMNS.AMOUNT];
    const currency = movement[COLUMNS.CURRENCY];
    
    // Skip if amount or currency is invalid
    if (!amount || !currency || isNaN(amount)) {
      Logger.log(`Movement ID ${movementId} has invalid amount (${amount}) or currency (${currency})`);
      return false;
    }
    
    // Get currency conversion service and attempt to fix the conversion
    const currencyConversionService = new CurrencyConversionService();
    
    const currencyValues = currencyConversionService.fixCurrencyConversion(
      amount, 
      currency
    );
    
    if (!currencyValues || (!currencyValues.clpValue && !currencyValues.usdValue && !currencyValues.gbpValue)) {
      Logger.log(`Failed to fix currency conversion for movement ID ${movementId}`);
      return false;
    }
    
    // Update the movement with the fixed currency values (formulas or calculated values)
    this.sheet.getRange(sheetRowIndex, COLUMNS.CLP_VALUE + 1).setValue(currencyValues.clpValue);
    this.sheet.getRange(sheetRowIndex, COLUMNS.USD_VALUE + 1).setValue(currencyValues.usdValue);
    this.sheet.getRange(sheetRowIndex, COLUMNS.GBP_VALUE + 1).setValue(currencyValues.gbpValue);
    
    Logger.log(`Fixed currency conversion for movement ID ${movementId}: CLP=${currencyValues.clpValue}, USD=${currencyValues.usdValue}, GBP=${currencyValues.gbpValue}`);
    return true;
  }

  /**
   * Fix currency conversions for all movements that have failed conversions
   * @returns {Object} Object with success count and failure count
   */
  fixAllFailedCurrencyConversions() {
    const failedMovements = this.getMovementsWithFailedCurrencyConversion();
    Logger.log(`Found ${failedMovements.length} movements with failed currency conversions`);
    
    let successCount = 0;
    let failureCount = 0;
    
    failedMovements.forEach(movement => {
      const movementId = movement[COLUMNS.ID];
      if (this.fixMovementCurrencyConversion(movementId)) {
        successCount++;
      } else {
        failureCount++;
      }
    });
    
    Logger.log(`Currency conversion fix completed: ${successCount} successful, ${failureCount} failed`);
    return { successCount, failureCount };
  }


}
