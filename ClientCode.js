function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expense Tracker')
    .addItem('Run Actions', 'main')
    .addToUi();
}


function main() {
  try {
    return ExpenseTrackerLibrary.main();
  } catch (error) {
    Logger.log(`Error in main workflow: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
    throw error;
  }
}
