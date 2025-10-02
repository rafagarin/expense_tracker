/**
 * Splitwise API service for fetching credit movements
 * Handles authentication and data retrieval from Splitwise API
 */

class SplitwiseService {
  constructor(clientProperties = null) {
    this.clientProperties = clientProperties;
    this.splitwiseApiKey = getApiKey(API_CONFIG.SPLITWISE.API_KEY_PROPERTY, clientProperties);
    this.splitwiseBaseUrl = API_CONFIG.SPLITWISE.BASE_URL;
    this.currentUserId = null;
    this.categoryService = new CategoryService();
  }

  /**
   * Get all credit movements (expenses paid by others) from Splitwise
   * @returns {Array} Array of credit movement objects
   */
  async getCreditMovements() {
    try {
      Logger.log('Fetching credit movements from Splitwise...');
      
      // Get current user ID first
      await this.getCurrentUserId();
      
      // Get all expenses from Splitwise
      const expenses = await this.getAllExpenses();
      
      // Filter for credit movements (expenses where I owe money)
      const creditMovements = this.filterCreditMovements(expenses);
      
      Logger.log(`Found ${creditMovements.length} credit movement(s) from Splitwise`);
      return creditMovements;
      
    } catch (error) {
      Logger.log(`Error fetching credit movements from Splitwise: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get all debit movements (expenses paid by me for others) from Splitwise
   * @returns {Array} Array of debit movement objects
   */
  async getDebitMovements() {
    try {
      Logger.log('Fetching debit movements from Splitwise...');
      
      // Get current user ID first
      await this.getCurrentUserId();
      
      // Get all expenses from Splitwise
      const expenses = await this.getAllExpenses();
      
      // Filter for debit movements (expenses where others owe me money)
      const debitMovements = this.filterDebitMovements(expenses);
      
      Logger.log(`Found ${debitMovements.length} debit movement(s) from Splitwise`);
      return debitMovements;
      
    } catch (error) {
      Logger.log(`Error fetching debit movements from Splitwise: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get the current user ID from Splitwise
   * @returns {Promise<number>} Current user ID
   */
  async getCurrentUserId() {
    if (this.currentUserId) {
      return this.currentUserId;
    }

    try {
      const url = `${this.splitwiseBaseUrl}/get_current_user`;
      
      const options = {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.splitwiseApiKey}`,
          'Content-Type': 'application/json'
        }
      };

      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        this.currentUserId = data.user.id;
        Logger.log(`Current user ID: ${this.currentUserId}`);
        return this.currentUserId;
      } else {
        throw new Error(`Failed to get current user: ${response.getResponseCode()}`);
      }
    } catch (error) {
      Logger.log(`Error getting current user ID: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get expenses from Splitwise API from the last 5 days with pagination
   * @returns {Array} Array of expenses from the last 5 days
   */
  async getAllExpenses() {
    const allExpenses = [];
    let offset = 0;
    const limit = 100; // Maximum allowed by API
    let hasMore = true;

    // Calculate date 5 days ago
    const fiveDaysAgo = new Date();
    fiveDaysAgo.setDate(fiveDaysAgo.getDate() - 5);
    const datedAfter = fiveDaysAgo.toISOString();

    Logger.log(`Fetching Splitwise expenses from ${datedAfter} onwards`);

    while (hasMore) {
      const url = `${this.splitwiseBaseUrl}/get_expenses?limit=${limit}&offset=${offset}&dated_after=${encodeURIComponent(datedAfter)}`;
      
      const options = {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.splitwiseApiKey}`,
          'Content-Type': 'application/json'
        }
      };

      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        Logger.log(`Splitwise API request failed with status: ${response.getResponseCode()}, response: ${errorText}`);
        throw new Error(`Splitwise API request failed with status: ${response.getResponseCode()}`);
      }

      const data = JSON.parse(response.getContentText());
      
      if (!data.expenses) {
        Logger.log('No expenses found in Splitwise response');
        break;
      }

      allExpenses.push(...data.expenses);
      
      // Check if we got fewer results than the limit, meaning we've reached the end
      if (data.expenses.length < limit) {
        hasMore = false;
      } else {
        offset += limit;
      }

      Logger.log(`Fetched ${data.expenses.length} expenses (total: ${allExpenses.length})`);
    }

    Logger.log(`Total expenses fetched from last 5 days: ${allExpenses.length}`);
    return allExpenses;
  }

  /**
   * Filter expenses to find credit movements (where I owe money)
   * @param {Array} expenses - All expenses from Splitwise
   * @returns {Array} Array of credit movements
   */
  filterCreditMovements(expenses) {
    const creditMovements = [];

    for (const expense of expenses) {
      Logger.log(`Processing expense ${expense.id}: ${expense.description}`);
      Logger.log(`Payment: ${expense.payment}, Repayments: ${expense.repayments ? expense.repayments.length : 0}, Deleted: ${expense.deleted_at}`);
      
      // Skip deleted expenses
      if (expense.deleted_at) {
        Logger.log(`Skipping expense ${expense.id} - it's deleted`);
        continue;
      }
      
      // Skip repayments (expenses marked as payments between users)
      if (expense.payment === true) {
        Logger.log(`Skipping expense ${expense.id} - it's a payment`);
        continue;
      }

      // Check if this expense has users where I owe money
      if (expense.users && expense.users.length > 0) {
        Logger.log(`Expense ${expense.id} has ${expense.users.length} users:`);
        expense.users.forEach((user, index) => {
          Logger.log(`  User ${index}: ID=${user.user_id}, owed_share=${user.owed_share}, paid_share=${user.paid_share}, name=${user.first_name} ${user.last_name}`);
        });
        
        // Find my share in this expense (where I owe money)
        const myShare = this.findMyShare(expense.users);
        Logger.log(`My share for expense ${expense.id}: ${myShare ? JSON.stringify(myShare) : 'null'}`);
        
        if (myShare && parseFloat(myShare.net_balance) < 0) {
          // This is a credit movement - someone paid for me and I owe them
          const amount = Math.abs(parseFloat(myShare.net_balance));
          Logger.log(`Found credit movement: ${expense.description}, amount: ${amount}`);
          creditMovements.push({
            splitwiseId: expense.id.toString(),
            description: expense.description,
            amount: amount,
            currency: expense.currency_code,
            date: expense.date,
            category: null, // Will be filled by AI later
            paidBy: this.getPaidByUser(expense),
            group: expense.group_id ? `Group ${expense.group_id}` : 'Personal',
            details: expense.details || null
          });
        }
      } else {
        Logger.log(`Expense ${expense.id} has no users array`);
      }
    }

    return creditMovements;
  }

  /**
   * Filter expenses to find debit movements (where others owe me money)
   * @param {Array} expenses - All expenses from Splitwise
   * @returns {Array} Array of debit movements
   */
  filterDebitMovements(expenses) {
    const debitMovements = [];

    for (const expense of expenses) {
      Logger.log(`Processing expense ${expense.id} for debit movements: ${expense.description}`);
      Logger.log(`Payment: ${expense.payment}, Deleted: ${expense.deleted_at}`);
      
      // Skip deleted expenses
      if (expense.deleted_at) {
        Logger.log(`Skipping expense ${expense.id} - it's deleted`);
        continue;
      }
      
      // Skip repayments (expenses marked as payments between users)
      if (expense.payment === true) {
        Logger.log(`Skipping expense ${expense.id} - it's a payment`);
        continue;
      }

      // Check if this expense has users where others owe me money
      if (expense.users && expense.users.length > 0) {
        Logger.log(`Expense ${expense.id} has ${expense.users.length} users for debit check:`);
        expense.users.forEach((user, index) => {
          Logger.log(`  User ${index}: ID=${user.user_id}, owed_share=${user.owed_share}, paid_share=${user.paid_share}, name=${user.first_name} ${user.last_name}`);
        });
        
        // Find my share in this expense (where I paid and others owe me)
        const myShare = this.findMyDebitShare(expense.users);
        Logger.log(`My debit share for expense ${expense.id}: ${myShare ? JSON.stringify(myShare) : 'null'}`);
        
        if (myShare && parseFloat(myShare.net_balance) > 0) {
          // This is a debit movement - I paid for others and they owe me
          const amount = parseFloat(myShare.net_balance);
          Logger.log(`Found debit movement: ${expense.description}, amount: ${amount}`);
          
          debitMovements.push({
            splitwiseId: expense.id.toString(),
            description: expense.description,
            amount: amount,
            currency: expense.currency_code,
            date: expense.date,
            category: null, // Will be filled by AI later
            owedBy: this.getOwedByUsers(expense.users),
            group: expense.group_id ? `Group ${expense.group_id}` : 'Personal',
            details: expense.details || null
          });
        }
      } else {
        Logger.log(`Expense ${expense.id} has no users array for debit check`);
      }
    }

    return debitMovements;
  }

  /**
   * Find my share in an expense (where I owe money)
   * @param {Array} users - Array of user share objects
   * @returns {Object|null} My share object or null if not found
   */
  findMyShare(users) {
    if (!this.currentUserId) {
      Logger.log('Current user ID not available, cannot identify my share');
      return null;
    }

    // Find the user share that belongs to me (current user ID)
    const myShare = users.find(user => user.user_id === this.currentUserId);
    
    if (myShare && myShare.net_balance && parseFloat(myShare.net_balance) < 0) {
      return myShare;
    }
    
    return null;
  }

  /**
   * Find my share in an expense (where others owe me money)
   * @param {Array} users - Array of user share objects
   * @returns {Object|null} My share object or null if not found
   */
  findMyDebitShare(users) {
    if (!this.currentUserId) {
      Logger.log('Current user ID not available, cannot identify my share');
      return null;
    }

    // Find the user share that belongs to me (current user ID)
    const myShare = users.find(user => user.user_id === this.currentUserId);
    
    if (myShare && myShare.net_balance && parseFloat(myShare.net_balance) > 0) {
      return myShare;
    }
    
    return null;
  }

  /**
   * Get the users who owe me money in this expense
   * @param {Array} users - Array of user share objects
   * @returns {string} Comma-separated list of users who owe money
   */
  getOwedByUsers(users) {
    if (!this.currentUserId) {
      return 'Unknown Users';
    }

    const owedByUsers = users
      .filter(user => user.user_id !== this.currentUserId && user.owed_share && parseFloat(user.owed_share) > 0)
      .map(user => {
        if (user.user) {
          const firstName = user.user.first_name || '';
          const lastName = user.user.last_name || '';
          return `${firstName} ${lastName}`.trim();
        }
        return 'Unknown User';
      })
      .filter(name => name !== '' && name !== 'Unknown User');

    return owedByUsers.length > 0 ? owedByUsers.join(', ') : 'Unknown Users';
  }

  /**
   * Get the user who paid for the expense
   * @param {Object} expense - The expense object from Splitwise
   * @returns {string} The name of the user who paid
   */
  getPaidByUser(expense) {
    if (expense.users && expense.users.length > 0) {
      // Find the user who paid (has a positive paid_share)
      const paidByUser = expense.users.find(user => user.paid_share && parseFloat(user.paid_share) > 0);
      if (paidByUser && paidByUser.user) {
        const firstName = paidByUser.user.first_name || '';
        const lastName = paidByUser.user.last_name || '';
        return `${firstName} ${lastName}`.trim() || 'Unknown User';
      }
    }
    return 'Unknown User';
  }

  /**
   * Map Splitwise category to our category system
   * @param {Object} splitwiseCategory - Category object from Splitwise
   * @returns {string} Our category name
   */
  mapSplitwiseCategory(splitwiseCategory) {
    if (!splitwiseCategory) {
      return this.getDefaultCategory();
    }

    const categoryName = splitwiseCategory.name.toLowerCase();
    const categories = this.categoryService.getCategories();
    
    // Map Splitwise categories to our dynamic categories
    const categoryMapping = {
      'food & dining': 'food',
      'transportation': 'transportation',
      'utilities': 'housing',
      'rent': 'housing',
      'healthcare': 'health',
      'entertainment': 'entertainment',
      'shopping': 'personal',
      'travel': 'transportation',
      'groceries': 'food',
      'restaurants': 'food',
      'gas': 'transportation',
      'public transportation': 'transportation',
      'medical': 'health',
      'pharmacy': 'health',
      'clothing': 'personal',
      'household': 'household',
      'work': 'work'
    };

    // Try exact match first
    if (categoryMapping[categoryName]) {
      const mappedCategory = categoryMapping[categoryName];
      if (this.categoryService.isValidCategory(mappedCategory)) {
        return mappedCategory;
      }
    }

    // Try partial matches
    for (const [key, value] of Object.entries(categoryMapping)) {
      if (categoryName.includes(key) || key.includes(categoryName)) {
        if (this.categoryService.isValidCategory(value)) {
          return value;
        }
      }
    }

    return this.getDefaultCategory();
  }

  /**
   * Get the default category (miscellaneous or first available)
   * @returns {string} Default category name
   */
  getDefaultCategory() {
    const categories = this.categoryService.getCategoryNames();
    
    // Try to find 'miscellaneous' first
    if (categories.includes('miscellaneous')) {
      return 'miscellaneous';
    }
    
    // Fall back to the first available category
    return categories[0] || 'miscellaneous';
  }

  /**
   * Get the configured Splitwise group ID
   * @returns {number} The group ID or default if not set
   */
  getSplitwiseGroupId() {
    const properties = PropertiesService.getScriptProperties();
    const groupId = properties.getProperty('SPLITWISE_GROUP_ID');
    return groupId ? parseInt(groupId) : SPLITWISE_CONFIG.DEFAULT_GROUP_ID;
  }

  /**
   * Get the configured other user ID for splitting
   * @returns {number|null} User ID of the person you split expenses with
   */
  getSplitwiseOtherUserId() {
    const properties = PropertiesService.getScriptProperties();
    const userId = properties.getProperty('SPLITWISE_OTHER_USER_ID');
    return userId ? parseInt(userId) : SPLITWISE_CONFIG.OTHER_USER_ID;
  }

  /**
   * Fetch and log all user IDs from Splitwise movements
   * This helps identify user IDs for proper expense splitting
   */
  async logSplitwiseUserIds() {
    try {
      Logger.log('Fetching Splitwise user IDs...');
      
      // Get current user ID first
      await this.getCurrentUserId();
      
      // Get all expenses from Splitwise
      const expenses = await this.getAllExpenses();
      
      if (expenses.length === 0) {
        Logger.log('No expenses found in Splitwise');
        return;
      }

      Logger.log(`Found ${expenses.length} expenses. Analyzing user IDs...`);
      
      const userIds = new Set();
      const userDetails = new Map();
      
      // Process each expense to extract user information
      for (const expense of expenses) {
        if (expense.users && Array.isArray(expense.users)) {
          for (const user of expense.users) {
            if (user.user && user.user.id) {
              const userId = user.user.id;
              const userName = user.user.first_name + ' ' + user.user.last_name;
              const userEmail = user.user.email;
              
              userIds.add(userId);
              
              // Store user details
              if (!userDetails.has(userId)) {
                userDetails.set(userId, {
                  id: userId,
                  name: userName,
                  email: userEmail,
                  isCurrentUser: userId === this.currentUserId
                });
              }
            }
          }
        }
      }
      
      // Log all user IDs and details
      Logger.log('=== SPLITWISE USER IDs ===');
      Logger.log(`Your User ID: ${this.currentUserId}`);
      Logger.log('Other Users:');
      
      for (const [userId, details] of userDetails) {
        if (!details.isCurrentUser) {
          Logger.log(`ID: ${userId} | Name: ${details.name} | Email: ${details.email}`);
        }
      }
      
      Logger.log('=== END USER IDs ===');
      Logger.log(`Total unique users found: ${userIds.size}`);
      
      return Array.from(userIds);
      
    } catch (error) {
      Logger.log(`Error fetching Splitwise user IDs: ${error.message}`);
      throw error;
    }
  }

  /**
   * Test the Splitwise API connection
   * @returns {boolean} True if connection is successful
   */
  async testConnection() {
    try {
      const userId = await this.getCurrentUserId();
      if (userId) {
        Logger.log(`Splitwise connection successful. User ID: ${userId}`);
        return true;
      } else {
        Logger.log('Splitwise connection failed: Could not get user ID');
        return false;
      }
    } catch (error) {
      Logger.log(`Splitwise connection test failed: ${error.message}`);
      return false;
    }
  }

  /**
   * Create an expense in Splitwise
   * @param {Object} expenseData - The expense data to create
   * @returns {string|null} The created expense ID or null if failed
   */
  async createExpense(expenseData) {
    try {
      const url = `${this.splitwiseBaseUrl}/create_expense`;
      
      // For debit movements, we need to show that others owe us money
      // The total cost is the amount we paid, but we only owe ourselves our portion
      const totalAmount = expenseData.amount;
      const personalAmount = expenseData.personalAmount || expenseData.amount;
      const othersAmount = totalAmount - personalAmount;
      
      Logger.log(`=== SPLITWISE DEBUG ===`);
      Logger.log(`Total Amount: ${totalAmount}`);
      Logger.log(`Personal Amount: ${personalAmount}`);
      Logger.log(`Others Amount: ${othersAmount}`);
      Logger.log(`Description: ${expenseData.description}`);
      Logger.log(`Other User ID: ${this.getSplitwiseOtherUserId()}`);
      Logger.log(`=======================`);
      
      // Build form-encoded payload according to API schema
      const payload = {
        cost: totalAmount.toString(),
        description: expenseData.description,
        date: expenseData.date,
        currency_code: expenseData.currency,
        group_id: expenseData.groupId || this.getSplitwiseGroupId(),
        split_equally: false, // We're providing explicit shares
        users__0__user_id: this.currentUserId,
        users__0__paid_share: totalAmount.toString(), // We paid the full amount
        users__0__owed_share: personalAmount.toString() // We only owe ourselves our portion
      };

      // Add other users if specified (for more complex splits)
      if (expenseData.otherUsers && expenseData.otherUsers.length > 0) {
        expenseData.otherUsers.forEach((user, index) => {
          payload[`users__${index + 1}__user_id`] = user.userId;
          payload[`users__${index + 1}__paid_share`] = '0';
          payload[`users__${index + 1}__owed_share`] = user.owedShare.toString();
        });
      } else if (othersAmount > 0) {
        // Use configured other user ID for splitting
        const otherUserId = this.getSplitwiseOtherUserId();
        if (otherUserId) {
          // The other user owes the full amount others owe
          payload['users__1__user_id'] = otherUserId;
          payload['users__1__paid_share'] = '0';
          payload['users__1__owed_share'] = othersAmount.toString();
          Logger.log(`Split ${othersAmount} with user ID: ${otherUserId}`);
        } else {
          Logger.log(`Note: No other user ID configured. Others owe: ${othersAmount}. Use setupSplitwiseOtherUserId() to configure user ID.`);
        }
      }
      
      Logger.log(`Final payload: ${JSON.stringify(payload)}`);

      const options = {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.splitwiseApiKey}`,
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: Object.keys(payload).map(key => `${key}=${encodeURIComponent(payload[key])}`).join('&')
      };

      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        Logger.log(`Splitwise create expense failed with status: ${response.getResponseCode()}, response: ${errorText}`);
        return null;
      }

      const responseData = JSON.parse(response.getContentText());
      
      // Check for errors in the response
      if (responseData.errors && responseData.errors.length > 0) {
        Logger.log(`Splitwise create expense errors: ${JSON.stringify(responseData.errors)}`);
        return null;
      }
      
      if (responseData.expenses && responseData.expenses.length > 0) {
        const expenseId = responseData.expenses[0].id;
        Logger.log(`Created Splitwise expense with ID: ${expenseId}`);
        return expenseId.toString();
      } else {
        Logger.log('No expense created in Splitwise response');
        return null;
      }

    } catch (error) {
      Logger.log(`Error creating Splitwise expense: ${error.message}`);
      return null;
    }
  }
}
