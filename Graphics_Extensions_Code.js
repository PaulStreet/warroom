function copyCompletes_() {
  copyCompletes_GFX();
  }

function copyCompletes_GFX() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Look at the active spreadsheet
  var sheet1 = ss.getSheetByName('New Request'); // Get the New Request sheet
  var sheet2 = ss.getSheetByName('Completed Request'); // Store in the completed sheet
  
  var data = sheet1.getRange(3,1, sheet1.getLastRow(), sheet1.getLastColumn()).getValues(); 
	// gets values of sheet starting at row 3
  var dest = []; //sets up an array
  for (var i = 0; i < data.length; i++ ) {
    Logger.log(data[i][34]); // Log all the data in the sheet
    if (data[i][4] !== '')  { //See if there is a gfx completion date
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
    if (row[4] !== '') {
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
  spreadsheet.addMenu('Automation', menuItems);
}
