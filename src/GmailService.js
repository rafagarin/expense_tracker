/**
 * Gmail service for searching and retrieving bank notification emails
 */

class GmailService {
  constructor(clientProperties = null) {
    this.clientProperties = clientProperties;
  }

  /**
   * Search for bank notification emails
   * @returns {Array} Array of Gmail thread objects
   */
  searchBankEmails() {
    const threads = GmailApp.search(GMAIL_QUERY);
    Logger.log(`Found ${threads.length} email thread(s) matching the query: ${GMAIL_QUERY}`);
    return threads;
  }

  /**
   * Get all messages from email threads
   * @param {Array} threads - Array of Gmail thread objects
   * @returns {Array} Array of Gmail message objects
   */
  getAllMessagesFromThreads(threads) {
    const messages = [];
    
    threads.forEach(thread => {
      const threadMessages = thread.getMessages();
      messages.push(...threadMessages);
    });
    
    Logger.log(`Found ${messages.length} total message(s) across all threads.`);
    return messages;
  }

  /**
   * Get messages that haven't been processed yet
   * @param {Set} existingGmailIds - Set of already processed Gmail IDs
   * @returns {Array} Array of unprocessed Gmail message objects
   */
  getUnprocessedMessages(existingGmailIds) {
    const threads = this.searchBankEmails();
    const allMessages = this.getAllMessagesFromThreads(threads);
    
    const unprocessedMessages = allMessages.filter(message => {
      const gmailId = message.getId();
      const isProcessed = existingGmailIds.has(gmailId);
      
      if (isProcessed) {
        Logger.log(`Skipping already processed email with ID: ${gmailId}`);
      }
      
      return !isProcessed;
    });
    
    Logger.log(`Found ${unprocessedMessages.length} unprocessed message(s).`);
    return unprocessedMessages;
  }
}
