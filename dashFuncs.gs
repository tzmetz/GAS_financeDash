// ----------------------------------------------------------------
// -------------------- DATA UPDATE SCRIPT ------------------------
// ----------------------------------------------------------------

// This script updates data to spreadsheet based on changes from source in drive 
// Once updated, new data is categorized by inputs from the user

// Apps Scripts Reference Documentation https://developers.google.com/apps-script/reference/spreadsheet
// Apps scripts tut https://zapier.com/learn/google-sheets/google-apps-script-tutorial/      https://spreadsheet.dev/how-to-import-csv-files-into-google-sheets-using-apps-script
// Java Script Reference: https://www.w3schools.com/js/js_arrays.asp
// Apps scripts basic tutorial: https://spreadsheet.dev/learn-coding-google-sheets-apps-script
// Methods for getting .lastRow(): https://yagisanatode.com/2019/05/11/google-apps-script-get-the-last-row-of-a-data-range-when-other-columns-have-content-like-hidden-formulas-and-check-boxes/
// V8 Runtime new additions: https://www.benlcollins.com/apps-script/apps-script-v8-runtime/

// TODO: dont run update if there are no changes
// TODO: dont write account # column  
// TODO: Automatic detection/categorization of chase credit card payment 
// TODO: automatic detection of transfer 
// TODOL automatic detection of loan payments

// Global Vars 
const CAT_COL = 9;
var CATS_LOCATIONS = "A9:A";

// ----------------------------- ON OPEN -----------------------------------
function onOpen(e) {
  
  // Create UI for importing data call
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Update Data")
  //.addItem("Import from URL","importCSVFromUrl")
  .addItem("Update Data From Drive", "importCSVFromDrive")
  .addToUi();
}

// -------------------------- IMPORTING DATA -------------------------------
  
function importCSVFromDrive() {
  // Declare UI Variables 
  var ui = SpreadsheetApp.getUi();
  
  // First Log Current State of the Data 
  var ss = SpreadsheetApp.getActive();
  var rawDataSheet = ss.getSheetByName("RawData");
  var lastR = rawDataSheet.getLastRow();
  
  var folderID = "1bw0UyVE2VqTamqir4GxrKxkyOvroQGLR"; 
  var files = findFilesInDrive(folderID);
  
  //Logic check on length of files
  if(files.length === 0) {
    displayToastAlert("Missing File in folder with ID: " + folderID);
    return; 
  } else if(files.length > 1) {
    displayToastAlert("Multiple files with fileName found in folder with ID: " + folderID);
    return;
  }
  
  var file = files[0];
  var data = Utilities.parseCsv(file.getBlob().getDataAsString());
  
  // Check For Differences Between New and Existing Data
  var diff = data.length - lastR; // Adding minus 1 to remove the header row
  
  // Now overwrite existing data with new data
  writeDataToSheet(data);
  
  if(diff > 0) {
    displayToastAlert(diff + " lines appended");
    // Shift Categories Column Down to Keep Link To Data by inserting blank cells 
    // for new categories starting at the top. Shift by number of new data rows
    rawDataSheet.getRange(2, 11, diff, 6).insertCells(SpreadsheetApp.Dimension.ROWS);
  } else if(diff < 0) {
    ui.alert("Lines Removed", ui.ButtonSet.OK);
    return;
  } else {
    displayToastAlert("No Change in Data");
  } 
}

// ---------------------------------- SUPPORT FUNCS -----------------------------

//Displays an Alert as a Toast message 
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, " Alert");
}

// Returns files in Google Drive that Have a Certain Name 
function findFilesInDrive(folderID) {
  //var files = DriveApp.getFilesByName(fileName);
  var files = DriveApp.getFolderById(folderID).getFilesByName("AccountHistory.csv");
  var result = [];
  
  // If files has length > 1
  while(files.hasNext())
    result.push(files.next()); //adds files.next() to result array https://www.w3schools.com/jsref/jsref_push.asp
  
  return result;
}

// Writes Data to a Sheet on Current Spreadsheet
function writeDataToSheet(data) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("RawData");
  
  // Clear sheet first (in case new data has fewer rows than old data)
  var lastR = sheet.getLastRow();
  sheet.getRange(1, 1, lastR, 7).clear();

  // Now write new data
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}


/**************************************************************************
 *
 * Writes a new user inputed category group to the category group database on FORMULAS2 sheet
 *
 */ 
function addNewGroupButton() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Input New Category Group', ui.ButtonSet.OK_CANCEL);
  
  if(response.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  
  // Check to see if this group already exists
  var formula2Sheet = SpreadsheetApp.getActive().getSheetByName("FORMULAS2");
  
  // First Get the last row in the groups DB that contains data 
  var searchRangeLastRow = formula2Sheet.getLastRow();
  var searchArray = formula2Sheet.getRange(2, 9, searchRangeLastRow).getValues();
  
  for (i = 0; i <= searchArray.length; i++) {
    if(searchArray[i].toString() == "") { // searchArray is an array containing multiple arrays, convert each array to a string to expose the empty arrays as emppty strings which can be checked for in the conditional
      var lastRow = i + 2; // Adding 2 to offset for header and the fact that array indexing starts at 0
      
      if(i == 0) { // warn user if DB needs to be initialized
        ui.alert("ERROR: Category Groups DB must be initialized with at least one group");
        return;
      }
      
      break;
    }
  }
  
  var groupDB = formula2Sheet.getRange(2, 9, lastRow-2).getValues();
  
  for (i = 0; i < groupDB.length; i++) {
    if( response.getResponseText() == groupDB[i].toString() )  {
      ui.alert("This group already exists");
      return;
    }
  }
  
  // Now write the user's input to the last row of the DB
  formula2Sheet.getRange(lastRow, 9).setValue(response.getResponseText());
}

/**************************************************************************
 *
 * Writes a new user inputed category tag to the category tag database on FORMULAS2 sheet
 *
 */ 
function addNewTagButton() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Input New Category Tag', ui.ButtonSet.OK_CANCEL);
  
  if(response.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  
  // Check to see if this tag already exists
  var formula2Sheet = SpreadsheetApp.getActive().getSheetByName("FORMULAS2");
  
  // First Get the last row in the groups DB that contains data 
  var searchRangeLastRow = formula2Sheet.getLastRow();
  var searchArray = formula2Sheet.getRange(2, 10, searchRangeLastRow).getValues();
  
  for (i = 0; i <= searchArray.length; i++) {
    if(searchArray[i].toString() == "") { // searchArray is an array containing multiple arrays, convert each array to a string to expose the empty arrays as emppty strings which can be checked for in the conditional
      var lastRow = i + 2; // Adding 2 to offset for header and the fact that array indexing starts at 0
      
      if(i == 0) { // warn user if DB needs to be initialized
        ui.alert("ERROR: Category Tags DB must be initialized with at least one group");
        return;
      }
      
      break;
    }
  }
  
  var tagDB = formula2Sheet.getRange(2, 10, lastRow-2).getValues();
  
  for (i = 0; i < tagDB.length; i++) {
    if( response.getResponseText() == tagDB[i].toString() )  {
      ui.alert("This group already exists");
      return;
    }
  }
  
  // Now write the user's input to the last row of the DB
  formula2Sheet.getRange(lastRow, 10).setValue(response.getResponseText());
}


/**************************************************************************
 *
 * Writes budgets user inputs to cells in "update budgets to" column on the dashboard
 *
 */ 
function updateBudgetsButton() {
  
  // Important Parameters 
  var dashCatTableSize = 20; // sets the max number of categories that the user can input
  var dashCatTableRowOffset = 9; // sets the row offset for where the cat table is on DashBoard sheet
  
  // Establish Sheets & ui
  var dashSheet = SpreadsheetApp.getActive().getSheetByName("DashBoard");
  var formulaSheet = SpreadsheetApp.getActive().getSheetByName("FORMULAS");
  var catSheet = SpreadsheetApp.getActive().getSheetByName("CATEGORIES");
  var ui = SpreadsheetApp.getUi();
  
  // First Make Sure That The User Is Not Overwriting Existing Budget Data From The Past 
  
  // Get current month and year and viewing month and year
  // Get What Month and Year the user is viewing in the dash sheet
  var viewingMonth = formulaSheet.getRange("C2").getValue();
  var viewingYear = dashSheet.getRange("I2").getValue();
  
  // Get What the current year and month are to compare these two 
  var yearToday = Number(Utilities.formatDate(new Date(), "GMT - 07:00", "yyyy"));
  var monthToday = Number(Utilities.formatDate(new Date(), "GMT - 07:00", "MM"));
  
  // If the user is about to overwrite old data, send an alert and stop execution
  if( (viewingMonth != monthToday && viewingYear != yearToday) || (viewingMonth != monthToday && viewingYear == yearToday) ) {
    ui.alert("Cannot Overwrite Existing Historical Data");
    return;
  }
  
  // Get Current Categories 
  var catsWithEmpties = dashSheet.getRange(dashCatTableRowOffset, 1, dashCatTableSize).getValues();
  
  // Remove all empty cells from cats array by loading all non empties to a new array
  var cats = [];
  for(var i = 0; i < catsWithEmpties.length; i++) {
    if(catsWithEmpties[i].toString() != "") { // finds an empty by converting array element to a string and compares to empty string
      cats.push(catsWithEmpties[i]);
    }
  }
  
  // Get What's Currently In The Update Budget To Column 
  var budgetsWithEmpties = dashSheet.getRange(dashCatTableRowOffset, 9, dashCatTableSize).getValues(); // this concatination and spreading turns the 2D array output from getValues into a 1D array
  
  // Remove all empty elements from budgets with empties, knowing that if there is no cat there cannot be a budget
  // First check to make sure that there were empties in the cat column in the first place
  if(cats.length < budgetsWithEmpties.length) {
    budgetsWithEmpties.splice( cats.length, (budgetsWithEmpties.length - cats.length)); // Use splice to remove all elements past the last category
  }
  
  // Check to see if there are any empties in budgetsWithEmpties. If there are, replace these with the current budget value
  var budgets = [];
  for(i = 0; i < budgetsWithEmpties.length; i++) {
    if(budgetsWithEmpties[i].toString() == "") {
      var defaultVal = dashSheet.getRange(i+dashCatTableRowOffset, 6).getValue() // If the user has left an empty cell in the update budget column then default to whatever the budget is currently set to
      // Check to make sure that there actually is a defaultVal available to write to the budgets array
      if(defaultVal.toString() == "") {
        budgets.push(0); // If there is no default val then set the budget to 0
      } 
      else {
        budgets.push(dashSheet.getRange(i+dashCatTableRowOffset, 6).getValue()); // If there is a default val then push that value
      }
    }
    else { // If element from budgetsWithEmpties is not empty then push that value to the new budgets array
      budgets.push(budgetsWithEmpties[i]);
    }
  }
 
  // Now we need to spread the budgets array to flatten all arrays within the array (from range.getValues()) resulting in a clean 1D array  
  budgets = [].concat(...budgets);
  
  // Now we need to transpose the budgets.lengthX1 array into a row array
  var budgetsTrans = [];
  budgetsTrans[0] = new Array(budgets.length); // Need to declare the first element of this array to be an empty row
  // Transpose algorithm. See here for a matrix transpose algorithm https://www.youtube.com/watch?v=HCtclMgx5VM
  for( i = 0; i < budgets.length; i++) {
    budgetsTrans[0][i] = budgets[i];
  }
  
  // Now Find Destination Row and Columns to Write New Budget Data 
  
  // Get Column of Months and Years in Category Sheet 
  // If viewing year = current year then we only need to look at a small number of rows in the cat table since the current year will be at the top of the table
  if(viewingYear == yearToday) {
    var catMonths = [].concat(...catSheet.getRange(5, 2, 12).getValues());
    var catYears = [].concat(...catSheet.getRange(5, 3, 12).getValues());
    
    // Find Destination Row
    for(i = 0; i < catMonths.length; i++) {
      if(catMonths[i] == viewingMonth && catYears[i] == viewingYear) {
        var destinationRow = i+5;
        break;
      }
    }
  } 
  else { // If we're not viewing the current year, we'll have to look through a larger dataset
    var lastRow = catSheet.getLastRow();
    var catMonths = [].concat(...catSheet.getRange(5, 2, lastRow-5).getValues());
    var catYears = [].concat(...catSheet.getRange(5, 3, lastRow-5).getValues());
    
    // Find Destination Row
    for(i = 0; i < catMonths.length; i++) {
      if(catMonths[i] == viewingMonth && catYears[i] == viewingYear) {
        var destinationRow = i+5;
        break;
      }
    }
  }  
  
  // Now Write to The Cat Table 
  catSheet.getRange(destinationRow, 4, 1, cats.length).setValues(budgetsTrans);
    
}  
  
  
