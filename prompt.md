This script helps me track my personal expenses. It:
- Reads expense info from Splitwise and from my bank's email notifications, with the help of AI
- Waits for me to provide a description of each movement along with additional instructions, and then parses this info with AI
- Adds movements to Splitwise
- Keeps track of loans and their payments

### Data model

The database columns are:
- timestamp (in ISO 8601 format, YYYY-MM-DDTHH:mm:ss.sssZ)
- direction (any of: Outflow, Inflow, Neutral)
- type (any of: Expense, Cash, Debit, Credit, Debit Repayment, Transfer)
- amount
- currency (I’m expecting CLP, USD, GBP)
- source_description (the usually non-descriptive string that the bank email provides, or the description from Splitwise)
- user_description (a user-provided description that will help with filling in more info about the movement, in step 3)
- comment (free text field, for instructions to AI)
- category (values are configured in the spreadsheet by the user)
- loan_status (for debit and credit movements: "Settled", "Pending Settlement", "Awaiting Splitwise Upload", "In Splitwise"; null for others)
- settled_movement_id (used for repayments, id of the movement that is fully or partially settled by this movement)
- clp_value (same as amount but converted to CLP)
- usd_value (same as amount but converted to USD)
- gbp_value (same as amount but converted to GBP)
- id (a unique id representing the movement)
- source ("gmail" or "accounting")
- gmail_id (ensures idempotency for movements that come from Gmail)
- accounting_system_id (ensures idempotency for movements that come from an external accounting system)

### Workflow steps

1. Entrypoint
   The script is triggered periodically. Each time it is run, it fetches info from two source types:
   A. It reads the email notifications from my bank. For each email that has not yet been added to the spreadsheet, it extracts the relevant info and saves it as a new row.
   B. It also reads movements from Splitwise.
   Additionally, one can also add movements manually, directly in the spreadsheet.
2. User input
   The user then provides additional information about the movement, which is saved in the "user_description" column, along with instructions for the system, which are saved in the "comment" column.
3. An AI reads the information entered by the user and performs relevant actions. This includes categorizing the movement and correctly formatting loans. See the examples section for more info.
4. In the future, the row can be again modified based on later movements such as loan repayments.


### Implementation
 
**Technology Stack:** Google Apps Script, connecting to Gmail, Google Sheets, Splitwise API, and Google AI Studio API.

**Project Structure:**
`Code.js` - Main entry point with public API functions and custom menu setup
`Config.js` - Configuration constants, database schema, and secure API key management
`Database.js` - Google Sheets database operations (CRUD, idempotency, batch processing)
`ExpenseTracker.js` - Core business logic orchestrating email processing, Splitwise integration, and movement creation
`GmailService.js` - Gmail integration for fetching and processing bank notification emails
`AIStudioService.js` - Google AI Studio integration for intelligent email parsing and category analysis using Gemini
`SplitwiseService.js` - Splitwise API integration for fetching credit and debit movements
`CategoryService.js` - Dynamic category management loading categories from Settings sheet
`CurrencyConversionService.js` - Currency conversion between CLP, USD, and GBP using rates from Values sheet
`appsscript.json` - Google Apps Script project configuration and permissions


### Examples

Here is how some example movements would be represented in the database:
1. If I bought something for myself, it is added as a movement with type: "Expense".
2. Cash withdrawals are added with type: "Cash", category: "miscellaneous".
3. Transfers between my own accounts will be added with direction: "Neutral", type: "Transfer".
4. When I lend money to people, it must be settled by a later movement. For example, say I paid a restaurant bill for multiple people but expect to be paid back what others spent. Then my part is added with type: "Expense", and the part of other people is added as type: "Debit".
  This debit can later be settled in two ways:
   1. With a bank transfer, which will be read from Gmail and saved with type "Debit Repayment".
   2. Or by adding it to Splitwise, which will also be saved with type "Debit Repayment".
5. When someone else lends me money, there is no way to detect the initial loan, so when I pay it with a bank transfer it will simply be added as a single movement with type: "Expense". Or it can be added to and read from Splitwise, in which case it will counted as a movement with type: "Credit".

In other words:
- Loan: someone owes someone else money.
- Debit: someone owes me money.
- Credit: I owe someone else money.
- Loans from Splitwise are added as a movement and that’s that.
- Credit loans from Gmail are simply counted as expenses.
- Debit loans from Gmail can get a little more complicated. It is possibly split from a parent movement, and then it might be added to Splitwise or it might be paid directly in a single or multiple payments.


Out of scope:

- Movements paid with credit card will just count as expenses (type: "Expense") right at that moment. They don't count as credit, and the actual payment of the credit card bill is later ignored.
- Cash withdrawals will just count as a miscellaneous expense.

### Setup

To set up the project:
- Copy the spreadsheet
- Copy the code
- Add values for `GOOGLE_AI_STUDIO_API_KEY`, `SPLITWISE_API_KEY`, `SPLITWISE_GROUP_ID`, `SPLITWISE_OTHER_USER_ID`
- Use [clasp](https://github.com/google/clasp) to push updates to the code

### Pending

1. Keeping track of settlements on Splitwise. Once loans (credit or debit) are marked as settled on Splitwise, they are marked as settled on the database. The system should show a summary of all unpaid loans.
2. Rules: I should be able to define rules for autocompletion of information (eg monthly payment of iCloud subscription is shared).
3. Figure out how to share it. This includes being very clear about what needs attention from the user.
4. Later versions can include a telegram bot UI, and more conversational capabilities.