function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Expense Tracker')
    .addItem('Run Actions', 'main')
    .addToUi();
}


function main() {
  try {
    const clientProperties = PropertiesService.getScriptProperties();
    return ExpenseTrackerLibrary.main(clientProperties);
  } catch (error) {
    Logger.log(`Error in main workflow: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
    throw error;
  }
}
