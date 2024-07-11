function mainFunction() {
    // manual conditions
    // pivotSheetNumber = 0
    // dataSheetNumber = 1
    // consultantSheetNumber = 2
    // collectedSheetNumber = 3
   
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
   
    for(let i = 4; i < sheets.length; i++){
      findAndReplace(sheets[i])
      fillAgentNames(sheets[i])
      updateDebtSheets(sheets[i], i)
      createPivotTables(sheets[i], i)
    }
  }
   
  function updateDebtSheets(activeSheet, i) {
    var sheet = activeSheet;
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
   
    var incomeSheet = sheets[3]; // Assuming Sheet4 is the income sheet
   
    const dataSheet = sheets[1];
    const dataValues = dataSheet.getDataRange().getValues();
   
    for (let i = 0; i < dataValues.length; i++) {
      if (dataValues[i][0] == activeSheet.getName() && dataValues[i][1] == 'updateDebtSheets') {
        Logger.log('updateDebtSheets has already run on ' + activeSheet.getName());
        return;
      }
    }
   
    // Get the data from the income sheet
    var incomeData = incomeSheet.getDataRange().getValues();
    var test = 0;
    // Extract dates from the sheet name
    if(i+1 >= sheets.length){
      var dates = [sheet.getName(), new Date()]
      test = 1;
    }
    else{
      var dates = [sheet.getName(), sheets[i+1].getName()]
      test = 2;
      }
   
    var startDate = new Date(dates[0]);
    var endDate = new Date(dates[1]);
   
    // Loop through the income data and filter between startDate and endDate
    for (var j = 1; j < incomeData.length; j++) { // Start from 1 to skip headers
      var incomeDate = incomeData[j][0];
      if (incomeDate >= startDate && incomeDate < endDate) {
        var invoiceNumber = incomeData[j][1];
        var amountPaid = incomeData[j][2];
   
        // Find the invoice number in the debt sheet and update column 22
        var invoiceRange = sheet.getRange(1, 6, sheet.getLastRow(), 1); // Column F (6th column)
        var invoiceValues = invoiceRange.getValues();
   
        for (var k = 0; k < invoiceValues.length; k++) {
          if (invoiceValues[k][0] == invoiceNumber) {
            var currentAmount = sheet.getRange(k + 1, 22).getValue(); // Column 22 (V)
            sheet.getRange(k + 1, 22).setValue(currentAmount + amountPaid);
            break;
          }
        }
      }
    }
    logScriptRun(activeSheet.getName(), 'updateDebtSheets');
  }
   
   
  function createPivotTables(activeSheet, i) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var destinationSheet = ss.getSheets()[0];
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
   
    const dataSheet = sheets[1];
    const dataValues = dataSheet.getDataRange().getValues();
   
    for (let i = 0; i < dataValues.length; i++) {
      if (dataValues[i][0] == activeSheet.getName() && dataValues[i][1] == 'createPivotTables') {
        Logger.log('createPivotTables has already run on ' + activeSheet.getName());
        return;
      }
    }
   
    // Clear the destination sheet
    // destinationSheet.clear();
   
    // Loop through sheets starting from sheet number 5
   
    var sheet = activeSheet
    var range = sheet.getDataRange();
   
    // Define the positions for the pivot tables
    var startRow = 1 + (i - 4) * 16; // Adjust vertical space between sets of pivot tables
    var startCol = 1;
   
    createPivotTable(sheet, range, destinationSheet, startRow, startCol, 12, 'SUM');
    createPivotTable(sheet, range, destinationSheet, startRow, startCol + 12, 6, 'COUNTUNIQUE');
    createPivotTable(sheet, range, destinationSheet, startRow, startCol + 24, 1, 'COUNTUNIQUE');
    createPivotTable(sheet, range, destinationSheet, startRow, startCol + 36, 22, 'SUM');
   
    logScriptRun(activeSheet.getName(), 'createPivotTables');
  }
   
  function createPivotTable(sourceSheet, sourceRange, destinationSheet, startRow, startCol, columnIndex, summarizeFunction) {
    const pivotTableRange = destinationSheet.getRange(startRow, startCol, 15, 16); // adjust range as needed
   
    let pivotTable = pivotTableRange.createPivotTable(sourceRange);
    pivotTable = pivotTableRange.createPivotTable(sourceRange);
   
    // Configure row and column groups
    pivotTable.addRowGroup(21); // 21st column "Consultant"
    pivotTable.addColumnGroup(13); // 13th column "VarstaFactura"
   
    // Add pivot values based on column name and summarize function
    pivotTable.addPivotValue(columnIndex, SpreadsheetApp.PivotTableSummarizeFunction[summarizeFunction]);
   
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
   
    const agentSheetData = sheets[2] // replace 'Sheet1' with the actual name if different
    var agents = [];
    const agentData = agentSheetData.getDataRange().getValues();
   
    for(let i = 0; i<agentSheetData.getLastColumn(); i++) {
      agents.push(agentData[0][i]);
    }
   
    var filterCriteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(agents);
   
    pivotTable.addFilter(21, filterCriteria)
  }
   
   
   
   
  function fillAgentNames(activeSheet) {
    // Open the spreadsheets
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const dataSheet = sheets[1];
    const dataValues = dataSheet.getDataRange().getValues();
   
    // Get the sheets
    const agentSheetData = sheets[2] // replace 'Sheet1' with the actual name if different
    const billingSheetData = activeSheet; // replace 'Sheet1' with the actual name if different
   
    for (let i = 0; i < dataValues.length; i++) {
      if (dataValues[i][0] == billingSheetData.getName() && dataValues[i][1] == 'fillAgentNames') {
        Logger.log('fillAgentNames has already run on ' + activeSheet.getName());
        return;
      }
    }
   
    // Get data from the agent sheet
    const agentData = agentSheetData.getDataRange().getValues();
    const agentMap = new Map();
   
    var numOfColumns = agentSheetData.getLastColumn()
   
    // Loop through the agent data and create a map of CC to agent names
    for (let i = 1; i < agentData.length; i++) { // assuming first row is header
      for (let j = 0; j < numOfColumns; j++) { // 8 columns of client codes
        const cc = agentData[i][j];
        if (cc) {
          const agentName = agentData[0][j]; // assuming agent names are in the first row
          agentMap.set(cc, agentName);
        }
      }
    }
   
    // Get data from the billing sheet
    const billingData = billingSheetData.getDataRange().getValues();
   
    // Loop through the billing data and fill in the agent names
    for (let i = 1; i < billingData.length; i++) { // assuming first row is header
      const cc = billingData[i][4]; // assuming CC is in the first column of billing sheet
      if (cc && agentMap.has(cc)) {
        billingData[i][20] = agentMap.get(cc); // assuming empty column for agent name is the second column
      }
    }
    // Write the updated data back to the billing sheet
    billingSheetData.getRange(1, 1, billingData.length, billingData[0].length).setValues(billingData);
   
    logScriptRun(billingSheetData.getName(), 'fillAgentNames');
  }
   
  function findAndReplace(activeSheet) {
    const sheet = activeSheet;
   
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const dataSheet = sheets[1];
    const dataValues = dataSheet.getDataRange().getValues();
   
    for (let i = 0; i < dataValues.length; i++) {
      if (dataValues[i][0] == sheet.getName() && dataValues[i][1] == 'findAndReplace') {
        Logger.log('findAndReplace has already run on ' + activeSheet.getName());
        return;
      }
    }
   
    // Define the find and replace pairs
    const findReplacePairs = [
      { find: '1-30 zile', replace: '1) 1-30 zile' },
      { find: '31-60 zile', replace: '2) 31-60 zile' },
      { find: '61-90 zile', replace: '3) 61-90 zile' },
      { find: '91-180 zile', replace: '4) 91-180 zile' },
      { find: '181-270 zile', replace: '5) 181-270 zile' },
      { find: '271-365 zile', replace: '6) 271-365 zile' },
      { find: '1-2 ani', replace: '7) 1-2 ani' },
      { find: '2-3 ani', replace: '8) 2-3 ani' },
      { find: '>3 ani', replace: '9) >3 ani' }
    ];
   
    // Loop through each find and replace pair
    findReplacePairs.forEach(pair => {
      sheet.createTextFinder(pair.find).replaceAllWith(pair.replace);
    });
   
    logScriptRun(sheet.getName(), 'findAndReplace');
  }
   
  function reverseFindAndReplace(activeSheet) {
    const sheet = activeSheet;
   
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const dataSheet = sheets[1];
    const dataValues = dataSheet.getDataRange().getValues();
   
   
    for (let i = 0; i < dataValues.length; i++) {
      if (dataValues[i][0] == sheet.getName() && dataValues[i][1] == 'reverseFindAndReplace') {
        Logger.log('reverseFindAndReplace has already run on ' + activeSheet.getName());
        return;
      }
    }
   
    // Define the find and replace pairs
    const findReplacePairs = [
      { find: '1) 1-30 zile', replace: '1-30 zile' },
      { find: '2) 31-60 zile', replace: '31-60 zile' },
      { find: '3) 61-90 zile', replace: '61-90 zile' },
      { find: '4) 91-180 zile', replace: '91-180 zile' },
      { find: '5) 181-270 zile', replace: '181-270 zile' },
      { find: '6) 271-365 zile', replace: '271-365 zile' },
      { find: '7) 1-2 ani', replace: '1-2 ani' },
      { find: '8) 2-3 ani', replace: '2-3 ani' },
      { find: '9) >3 ani', replace: '>3 ani' }
    ];
   
    // Loop through each find and replace pair
    findReplacePairs.forEach(pair => {
      sheet.createTextFinder(pair.find).replaceAllWith(pair.replace);
    });
   
    logScriptRun(sheet.getName(), 'reverseFindAndReplace');
  }
   
  function logScriptRun(sheetName, scriptType) {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const dataSheet = sheets[1];
   
    const lastRow = dataSheet.getLastRow();
    dataSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[sheetName, scriptType]]);
  }