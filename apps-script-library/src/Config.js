/**
 * Configuration constants for the expense tracker
 */

// Spreadsheet configuration
const SPREADSHEET_TAB_NAME = 'Movements';

// Gmail search configuration
const GMAIL_QUERY = '(label:expenses newer_than:8d) OR label:expenses/manual ';

// Database column indices (0-based)
const COLUMNS = {
  TIMESTAMP: 0,
  DIRECTION: 1,
  TYPE: 2,
  AMOUNT: 3,
  CURRENCY: 4,
  SOURCE_DESCRIPTION: 5,
  USER_DESCRIPTION: 6,
  COMMENT: 7,
  AI_COMMENT: 8,
  CATEGORY: 9,
  STATUS: 10,
  SETTLED_MOVEMENT_ID: 11,
  CLP_VALUE: 12,
  USD_VALUE: 13,
  GBP_VALUE: 14,
  ORIGINAL_AMOUNT: 15,
  ID: 16,
  SOURCE: 17,
  SOURCE_ID: 18,
  ACCOUNTING_SYSTEM_ID: 19
};

// Supported currencies
const CURRENCIES = {
  CLP: 'CLP',
  USD: 'USD',
  GBP: 'GBP'
};

// Movement types
const MOVEMENT_TYPES = {
  EXPENSE: 'Expense',
  CASH: 'Cash',
  DEBIT: 'Debit',
  CREDIT: 'Credit',
  DEBIT_REPAYMENT: 'Debit Repayment'
};

// Movement directions
const DIRECTIONS = {
  OUTFLOW: 'Outflow',
  INFLOW: 'Inflow',
  NEUTRAL: 'Neutral'
};


// Status values
const STATUS = {
  SETTLED: 'Settled',
  PENDING_DIRECT_SETTLEMENT: 'Pending Settlement',
  PENDING_SPLITWISE_SETTLEMENT: 'Awaiting Splitwise Upload',
  IN_SPLITWISE: 'In Splitwise'
};


// Source values
const SOURCES = {
  GMAIL: 'gmail',
  MONZO: 'monzo',
  SPLITWISE: 'splitwise'
};

// Splitwise configuration
const SPLITWISE_CONFIG = {
  DEFAULT_GROUP_ID: 0, // Set this to your default group ID, or 0 for personal expenses
  OTHER_USER_ID: null // User ID of the person you commonly split expenses with
};

// API Configuration
const API_CONFIG = {
  // Google AI Studio API configuration
  GOOGLE_AI_STUDIO: {
    BASE_URL: 'https://generativelanguage.googleapis.com/v1beta',
    MODEL: 'gemini-2.0-flash-lite',
    API_KEY_PROPERTY: 'GOOGLE_AI_STUDIO_API_KEY'
  },
  // Splitwise API configuration
  SPLITWISE: {
    BASE_URL: 'https://secure.splitwise.com/api/v3.0',
    API_KEY_PROPERTY: 'SPLITWISE_API_KEY'
  },
  // Monzo API configuration
  MONZO: {
    BASE_URL: 'https://api.monzo.com',
    ACCESS_TOKEN_PROPERTY: 'MONZO_ACCESS_TOKEN',
    REFRESH_TOKEN_PROPERTY: 'MONZO_REFRESH_TOKEN',
    CLIENT_ID_PROPERTY: 'MONZO_CLIENT_ID',
    CLIENT_SECRET_PROPERTY: 'MONZO_CLIENT_SECRET'
  }
};

/**
 * Get API key from PropertiesService
 * This is the secure way to store API keys in Google Apps Script
 * @param {string} keyName - The property name for the API key
 * @param {Object} clientProperties - Optional client properties object
 * @returns {string} The API key
 */
function getApiKey(keyName, clientProperties = null) {
  let properties;
  
  if (clientProperties) {
    // Use client properties if provided
    properties = clientProperties;
  } else {
    // Fall back to script properties (for backward compatibility)
    properties = PropertiesService.getScriptProperties();
  }
  
  const apiKey = properties.getProperty(keyName);
  
  if (!apiKey) {
    throw new Error(`API key '${keyName}' not found. Please set it in the script properties.`);
  }
  
  return apiKey;
}

/**
 * Set API key in PropertiesService
 * This should be called once to store your API key securely
 * @param {string} keyName - The property name for the API key
 * @param {string} apiKey - The API key value
 */
function setApiKey(keyName, apiKey) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(keyName, apiKey);
  Logger.log(`API key '${keyName}' has been stored securely.`);
}

// Configuration is now available globally in Google Apps Script
