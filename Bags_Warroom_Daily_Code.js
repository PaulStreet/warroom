function copyCompletes_() {
  copyCompletes_LI();
  copyCompletes_CPO();
  copyCompletes_EI();
  copyCompletes_SI();
  }

function copyCompletes_LI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Look at the active spreadsheet
  var sheet1 = ss.getSheetByName('Load_Issues'); // Get the Load sheet
  var sheet2 = ss.getSheetByName('Completed_LO'); // Store the completed sheet
  
  var data = sheet1.getRange(3,1, sheet1.getLastRow(), sheet1.getLastColumn()).getValues(); 
	// gets values of sheet starting at row 3
  var dest = []; //sets up an array
  for (var i = 0; i < data.length; i++ ) {
    Logger.log(data[i][9]); // Log all the data in the sheet
    if (data[i][3] !== '')  { //See if there is a completion date
      dest.push(data[i]); // store data in a dest array
    }
  } // here is the end of the for loop

  Logger.log(dest) ; // log the dest array

  if (dest.length > 0 ) { // if dest array has values write it the Completed sheet
    sheet2.getRange(sheet2.getLastRow()+1,1,dest.length,dest[0].length).setValues(dest);
  }
  
  var rows = sheet1.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var rowsDeleted = 0;
  for (var i = 2; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[3] !== '') {
      sheet1.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
}

function copyCompletes_CPO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Critical_Problem_Orders');
  var sheet2 = ss.getSheetByName('Completed_CPO');
  
  var data = sheet1.getRange(3,1, sheet1.getLastRow(), sheet1.getLastColumn()).getValues(); 
	// gets values of sheet starting at row 3
  var dest = []; //sets up an array
  for (var i = 0; i < data.length; i++ ) {
    Logger.log(data[i][12]); // Log all the data in the sheet
    if (data[i][2] !== '')  { //See if there is a completion date
      dest.push(data[i]); // store data in an array
    }
  } // here is the end of the for loop

  Logger.log(dest) ; // log the dest array instead

  if (dest.length > 0 ) { // if array has values write it the Completed sheet
    sheet2.getRange(sheet2.getLastRow()+1,1,dest.length,dest[0].length).setValues(dest);
  }
  
  var rows = sheet1.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var rowsDeleted = 0;
  for (var i = 2; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[2] !== '') {
      sheet1.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
}

function copyCompletes_EI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Equipment_Issues');
  var sheet2 = ss.getSheetByName('Completed_EI');
  
  var data = sheet1.getRange(3,1, sheet1.getLastRow(), sheet1.getLastColumn()).getValues(); 
	// gets values of sheet starting at row 3
  var dest = []; //sets up an array
  for (var i = 0; i < data.length; i++ ) {
    Logger.log(data[i][11]); // Log all the data in the sheet
    if (data[i][3] !== '')  { //See if there is a completion date
      dest.push(data[i]); // store data in an array
    }
  } // here is the end of the for loop

  Logger.log(dest) ; // log the dest array instead

  if (dest.length > 0 ) { // if array has values write it the Completed sheet
    sheet2.getRange(sheet2.getLastRow()+1,1,dest.length,dest[0].length).setValues(dest);
  }
  
  var rows = sheet1.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var rowsDeleted = 0;
  for (var i = 2; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[3] !== '') {
      sheet1.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
}

function copyCompletes_SI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('Supply_Issues');
  var sheet2 = ss.getSheetByName('Completed_SI');
  
  var data = sheet1.getRange(3,1, sheet1.getLastRow(), sheet1.getLastColumn()).getValues(); 
	// gets values of sheet starting at row 3
  var dest = []; //sets up an array
  for (var i = 0; i < data.length; i++ ) {
    Logger.log(data[i][11]); // Log all the data in the sheet
    if (data[i][3] !== '')  { //See if there is a completion date
      dest.push(data[i]); // store data in an array
    }
  } // here is the end of the for loop

  Logger.log(dest) ; // log the dest array instead

  if (dest.length > 0 ) { // if array has values write it the Completed sheet
    sheet2.getRange(sheet2.getLastRow()+1,1,dest.length,dest[0].length).setValues(dest);
  }
  
  var rows = sheet1.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var rowsDeleted = 0;
  for (var i = 2; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[3] !== '') {
      sheet1.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Process Completes', functionName: 'copyCompletes_'}
  ];
  spreadsheet.addMenu('Metrics', menuItems);
}
