/**
 * Google AI Studio service for enhanced email parsing using Gemini
 */

class GoogleAIStudioService {
  constructor(clientProperties = null) {
    this.clientProperties = clientProperties;
    this.googleAIStudioApiKey = getApiKey(API_CONFIG.GOOGLE_AI_STUDIO.API_KEY_PROPERTY, clientProperties);
    this.googleAIStudioBaseUrl = API_CONFIG.GOOGLE_AI_STUDIO.BASE_URL;
    this.googleAIStudioModel = API_CONFIG.GOOGLE_AI_STUDIO.MODEL;
    this.categoryService = new CategoryService();
  }

  /**
   * Parse email content using Google AI Studio to extract transaction information
   * @param {string} emailBody - The email body text
   * @param {string} gmailId - The Gmail message ID
   * @returns {Object|null} Parsed transaction data or null if parsing fails
   */
  async parseEmailWithGoogleAIStudio(emailBody, gmailId) {
    try {
      const prompt = this.createGoogleAIStudioParsingPrompt(emailBody);
      const response = await this.callGoogleAIStudioAPI(prompt);
      
      if (response && response.candidates && response.candidates.length > 0) {
        return this.parseGoogleAIStudioResponse(response, gmailId);
      }
      
      return null;
    } catch (error) {
      Logger.log(`Error parsing email with Google AI Studio: ${error.message}`);
      return null;
    }
  }

  /**
   * Create a prompt for Google AI Studio to parse bank email content
   * @param {string} emailBody - The email body text
   * @returns {string} The formatted prompt
   */
  createGoogleAIStudioParsingPrompt(emailBody) {
    return `You are an expert at parsing bank transaction emails for expense tracking. Extract transaction information from the following email and return it in JSON format.

Email content:
${emailBody}

IMPORTANT: Classify the transaction type using ONLY these specific values:
- "expense": Regular purchases/expenses paid by me (most common)
- "cash": Cash withdrawals from ATM or bank
- "debit": Money I lent to other people (I paid but expect to be paid back)
- "credit": Money someone else lent to me (they paid for me, I owe them)
- "debit repayment": Someone paying me back money I lent them

Please extract the following information and return it as a JSON object:
- amount: The transaction amount as a number
- currency: The currency code (CLP, USD, or GBP)
- source_description: The merchant/description from the email
- timestamp: The transaction timestamp in ISO 8601 format (YYYY-MM-DDTHH:mm:ss.sssZ)
- transaction_type: ONE of the 5 types listed above (expense, cash, debit, credit, debit repayment)

Rules for classification:
- Regular purchases (food, gas, shopping, etc.) = "expense"
- ATM withdrawals = "cash"
- If I paid for a group but others will pay me back = "debit"
- If someone else paid for me and I need to pay them back = "credit"
- Bank transfers that are clearly repayments = "debit repayment"

If you cannot extract any of these fields, set them to null. Only return valid JSON, no additional text.`;
  }

  /**
   * Call the Google AI Studio API with the given prompt
   * @param {string} prompt - The prompt to send to the API
   * @returns {Object} The API response
   */
  async callGoogleAIStudioAPI(prompt) {
    const url = `${this.googleAIStudioBaseUrl}/models/${this.googleAIStudioModel}:generateContent`;
    
    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }]
    };

    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-goog-api-key': this.googleAIStudioApiKey
      },
      payload: JSON.stringify(payload)
    };

    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      const errorText = response.getContentText();
      Logger.log(`Google AI Studio API request failed with status: ${response.getResponseCode()}, response: ${errorText}`);
      throw new Error(`Google AI Studio API request failed with status: ${response.getResponseCode()}`);
    }

    return JSON.parse(response.getContentText());
  }

  /**
   * Parse the Google AI Studio response and convert it to our transaction format
   * @param {Object} apiResponse - The full API response object
   * @param {string} gmailId - The Gmail message ID
   * @returns {Object|null} Parsed transaction object or null if parsing fails
   */
  parseGoogleAIStudioResponse(apiResponse, gmailId) {
    try {
      // Extract the text content from the Google AI Studio API response
      if (!apiResponse.candidates || apiResponse.candidates.length === 0) {
        Logger.log('No candidates found in Google AI Studio response');
        return null;
      }

      const textContent = apiResponse.candidates[0].content.parts[0].text;
      
      // Try to parse the text content as JSON
      let parsedData;
      try {
        // Clean the response to extract JSON
        const jsonMatch = textContent.match(/\{[\s\S]*\}/);
        if (!jsonMatch) {
          Logger.log('No JSON found in Google AI Studio response text');
          return null;
        }
        parsedData = JSON.parse(jsonMatch[0]);
      } catch (jsonError) {
        Logger.log(`Failed to parse JSON from Google AI Studio response: ${jsonError.message}`);
        Logger.log(`Google AI Studio response text: ${textContent}`);
        return null;
      }
      
      // Validate required fields
      if (!parsedData.amount || !parsedData.currency) {
        Logger.log('Missing required fields in Google AI Studio response');
        Logger.log(`Parsed data: ${JSON.stringify(parsedData)}`);
        return null;
      }

      // Map AI transaction_type to our movement type constants
      const movementType = this.mapTransactionTypeToMovementType(parsedData.transaction_type);

      return {
        gmailId,
        timestamp: parsedData.timestamp || new Date().toISOString(),
        amount: parseFloat(parsedData.amount),
        currency: parsedData.currency.toUpperCase(),
        sourceDescription: parsedData.source_description || 'AI Parsed Transaction',
        transactionType: movementType
      };
    } catch (error) {
      Logger.log(`Error parsing Google AI Studio response: ${error.message}`);
      return null;
    }
  }

  /**
   * Map AI transaction type to our movement type constants
   * @param {string} transactionType - The transaction type from AI
   * @returns {string} The corresponding movement type constant
   */
  mapTransactionTypeToMovementType(transactionType) {
    if (!transactionType) {
      return MOVEMENT_TYPES.EXPENSE; // Default fallback
    }

    const typeMapping = {
      'expense': MOVEMENT_TYPES.EXPENSE,
      'cash': MOVEMENT_TYPES.CASH,
      'debit': MOVEMENT_TYPES.DEBIT,
      'credit': MOVEMENT_TYPES.CREDIT,
      'debit repayment': MOVEMENT_TYPES.DEBIT_REPAYMENT
    };

    return typeMapping[transactionType.toLowerCase()] || MOVEMENT_TYPES.EXPENSE;
  }

  /**
   * Analyze user description and determine the appropriate category and split information
   * @param {string} userDescription - The user-provided description
   * @param {Object} movementData - Additional movement context (amount, source_description, etc.)
   * @returns {Object|null} Analysis result with category and split information or null if analysis fails
   */
  async analyzeCategory(userDescription, movementData = {}) {
    try {
      const prompt = this.createCategoryAnalysisPrompt(userDescription, movementData);
      const response = await this.callGoogleAIStudioAPI(prompt);
      
      if (response && response.candidates && response.candidates.length > 0) {
        return this.parseCategoryAnalysisResponse(response);
      }
      
      return null;
    } catch (error) {
      Logger.log(`Error analyzing category: ${error.message}`);
      return null;
    }
  }

  /**
   * Create a prompt for Google AI Studio to analyze and categorize user descriptions
   * @param {string} userDescription - The user-provided description
   * @param {Object} movementData - Additional movement context
   * @returns {string} The formatted prompt
   */
  createCategoryAnalysisPrompt(userDescription, movementData) {
    const { amount, currency, sourceDescription, type, direction } = movementData;
    const categoriesList = this.categoryService.getCategoriesForAIPrompt();
    
    return `You are an expert at categorizing personal expenses and determining if they need to be split. Analyze the following user description and determine the most appropriate category and split information.

User Description: "${userDescription}"

Note: The description above may include both the main description and any comments/instructions. Look for split instructions in both parts.

Additional Context:
- Amount: ${amount ? `${currency} ${amount}` : 'Not provided'}
- Source Description: ${sourceDescription || 'Not provided'}
- Type: ${type || 'Not provided'}
- Direction: ${direction || 'Not provided'}

Available Categories (choose ONLY one):
${categoriesList}

Split Analysis:
Determine if this expense should be split into two parts. Look for indicators in both the main description and any comments/instructions:
- "split with", "shared with", "paid for group", "dinner with friends"
- "my part", "their part", "half", "portion"
- "dividir", "compartir", "mitad" (Spanish split indicators)
- References to other people or splitting costs
- Group activities where only part of the cost is personal
- Instructions in comments like "dividir" or "split"

Rules for categorization:
1. Choose the category that best represents the primary purpose of the expense
2. If it could fit multiple categories, choose the most specific one
3. Consider the context and amount when making decisions
4. When in doubt, choose "miscellaneous"

Return a JSON object with the following structure:
{
  "category": "category_name",
  "needs_split": true/false,
  "split_amount": number (only if needs_split is true, the amount for the personal portion),
  "split_category": "category_name" (only if needs_split is true, category for the personal portion),
  "clean_description": "clean description without split instructions",
  "split_instructions": "split instructions or comments"
}

If needs_split is true, split_amount should be the amount that represents the user's personal portion of the expense.
The clean_description should be the main description without any split-related instructions.
The split_instructions should contain any split-related instructions or comments.`;
  }

  /**
   * Parse the Google AI Studio response for category analysis
   * @param {Object} apiResponse - The full API response object
   * @returns {Object|null} Analysis result with category and split information or null if parsing fails
   */
  parseCategoryAnalysisResponse(apiResponse) {
    try {
      if (!apiResponse.candidates || apiResponse.candidates.length === 0) {
        Logger.log('No candidates found in category analysis response');
        return null;
      }

      const textContent = apiResponse.candidates[0].content.parts[0].text.trim();
      
      // Try to parse the text content as JSON
      let parsedData;
      try {
        // Clean the response to extract JSON
        const jsonMatch = textContent.match(/\{[\s\S]*\}/);
        if (!jsonMatch) {
          Logger.log('No JSON found in category analysis response text');
          return null;
        }
        parsedData = JSON.parse(jsonMatch[0]);
      } catch (jsonError) {
        Logger.log(`Failed to parse JSON from category analysis response: ${jsonError.message}`);
        Logger.log(`Category analysis response text: ${textContent}`);
        return null;
      }
      
      // Validate required fields
      if (!parsedData.category) {
        Logger.log('Missing required category field in analysis response');
        Logger.log(`Parsed data: ${JSON.stringify(parsedData)}`);
        return null;
      }

      // Validate that the category is one of our valid categories
      const validCategories = this.categoryService.getCategoryNames();
      if (!validCategories.includes(parsedData.category)) {
        Logger.log(`Invalid category in response: ${parsedData.category}`);
        return null;
      }

      // If needs_split is true, validate split fields
      if (parsedData.needs_split === true) {
        if (typeof parsedData.split_amount !== 'number' || !parsedData.split_category) {
          Logger.log('Invalid split data in response');
          Logger.log(`Parsed data: ${JSON.stringify(parsedData)}`);
          return null;
        }

        // Validate split category
        if (!validCategories.includes(parsedData.split_category)) {
          Logger.log(`Invalid split category in response: ${parsedData.split_category}`);
          return null;
        }
      }

      return {
        category: parsedData.category,
        needs_split: parsedData.needs_split === true,
        split_amount: parsedData.split_amount || null,
        split_category: parsedData.split_category || null,
        clean_description: parsedData.clean_description || parsedData.category,
        split_instructions: parsedData.split_instructions || null
      };
    } catch (error) {
      Logger.log(`Error parsing category analysis response: ${error.message}`);
      return null;
    }
  }

  /**
   * Match a debit repayment to a pending debit movement using AI
   * @param {Object} repaymentMovement - The debit repayment movement
   * @param {Array} pendingDebits - Array of pending debit movements
   * @returns {number|null} The ID of the matching debit movement or null if no match found
   */
  async matchDebitRepayment(repaymentMovement, pendingDebits) {
    try {
      if (!pendingDebits || pendingDebits.length === 0) {
        Logger.log('No pending debit movements to match against');
        return null;
      }

      const prompt = this.createDebitMatchingPrompt(repaymentMovement, pendingDebits);
      const response = await this.callGoogleAIStudioAPI(prompt);
      
      if (response && response.candidates && response.candidates.length > 0) {
        return this.parseDebitMatchingResponse(response);
      }
      
      return null;
    } catch (error) {
      Logger.log(`Error matching debit repayment: ${error.message}`);
      return null;
    }
  }

  /**
   * Create a prompt for Google AI Studio to match debit repayments to pending debits
   * @param {Object} repaymentMovement - The debit repayment movement
   * @param {Array} pendingDebits - Array of pending debit movements
   * @returns {string} The formatted prompt
   */
  createDebitMatchingPrompt(repaymentMovement, pendingDebits) {
    const { amount, currency, userDescription, comment, timestamp } = repaymentMovement;
    
    // Format pending debits for the prompt
    const pendingDebitsList = pendingDebits.map(debit => {
      const debitDate = new Date(debit[COLUMNS.TIMESTAMP]).toLocaleDateString();
      return `ID: ${debit[COLUMNS.ID]} | Date: ${debitDate} | Amount: ${debit[COLUMNS.CURRENCY]} ${debit[COLUMNS.AMOUNT]} | Description: "${debit[COLUMNS.USER_DESCRIPTION] || debit[COLUMNS.SOURCE_DESCRIPTION]}" | Comment: "${debit[COLUMNS.COMMENT] || 'None'}"`;
    }).join('\n');

    return `You are an expert at matching financial transactions. I need you to match a debit repayment to one of the pending debit movements.

REPAYMENT MOVEMENT TO MATCH:
- Amount: ${currency} ${amount}
- Description: "${userDescription || 'None'}"
- Comment: "${comment || 'None'}"
- Date: ${new Date(timestamp).toLocaleDateString()}

PENDING DEBIT MOVEMENTS:
${pendingDebitsList}

INSTRUCTIONS:
1. Analyze the repayment description and comment to understand what it's paying for
2. Look for matching patterns with the pending debit movements:
   - Similar amounts (exact match preferred, but small differences acceptable)
   - Related descriptions (e.g., "lunch" matches "dinner", "coffee" matches "cafe")
   - Time proximity (repayments usually happen within a few days/weeks of the original debit)
   - Context clues in comments (e.g., "Payment for this week's lunch" should match a lunch-related debit)

3. IMPORTANT: The repayment must be at least 95% of the debit amount to be considered a valid settlement. Prioritize matches where:
   - The repayment amount is close to or equal to the debit amount
   - The repayment amount is at least 95% of the debit amount
   - If multiple matches are possible, choose the one with the closest amount match

4. Consider these matching criteria in order of importance:
   - Amount similarity (exact match preferred, repayment should be ≥95% of debit amount)
   - Description similarity (keywords, context, activity type)
   - Time proximity (more recent debits are more likely)
   - Comment context (explicit references to specific activities)

5. Return ONLY the ID of the best matching debit movement, or "null" if no good match is found.
   - Only return a match if the repayment amount is at least 95% of the debit amount
   - If no match meets the 95% threshold, return "null"

Return format: Just the ID number (e.g., "123") or "null" if no match found.`;
  }

  /**
   * Parse the Google AI Studio response for debit matching
   * @param {Object} apiResponse - The full API response object
   * @returns {number|null} The ID of the matching debit movement or null if no match found
   */
  parseDebitMatchingResponse(apiResponse) {
    try {
      if (!apiResponse.candidates || apiResponse.candidates.length === 0) {
        Logger.log('No candidates found in debit matching response');
        return null;
      }

      const textContent = apiResponse.candidates[0].content.parts[0].text.trim();
      
      // Try to extract a number (the ID) or "null"
      const idMatch = textContent.match(/(\d+)/);
      if (idMatch) {
        const matchedId = parseInt(idMatch[1], 10);
        Logger.log(`AI matched debit repayment to movement ID: ${matchedId}`);
        return matchedId;
      }
      
      // Check for "null" response
      if (textContent.toLowerCase().includes('null')) {
        Logger.log('AI found no matching debit movement');
        return null;
      }
      
      Logger.log(`Unexpected debit matching response: ${textContent}`);
      return null;
    } catch (error) {
      Logger.log(`Error parsing debit matching response: ${error.message}`);
      return null;
    }
  }

}
