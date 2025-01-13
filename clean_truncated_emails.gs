function cleanUp(){
/**THE MAIN GOAL OF THIS MACRO IS TO CLEAN UP THE DUMMY DATA FROM MANUAL COPYING AND PASTING OF MODULE PROGRESS FROM ADMIN PORTAL. */

/** assigning app(variable) to an Object(Spreadsheet App)*/
  var app = SpreadsheetApp;

/** retrieve the active spreadsheet we working on, make sure you enter the active spreadsheet*/  
  var ss = app.getActiveSpreadsheet();

/** retrive the tab we working on*/  
  var activeTab = ss.getActiveSheet();

/**retrieve the last row 
 * activeTab.getRange(X,Y,Z,A).activate();
 * X--starting row index
 * Y--starting column index
 * Z--range of rows
 * A--range of columns
 * activeTab.getRange(1,1,1,1).activate(); -- This will redirect you to the starting cell
 * activeTab.getRange(1,1,1,1).getValue(); -- This will get the value of the cell A1
 * activeTab.getRange(2,5,1,1).getValue(); -- This will get the value of row 2, column E
 * Logger.log(activeTab.getRange(2,5,1,1).getValue());
*

/**Calculate the Total unique users, by getting the last row and dividing by 6 since each user has 6 unique data fields.
 * Then round off using toFixed()
 * Logger.log(ar)
 */
  var lastRow = activeTab.getLastRow();
  var totalUniqueUsers = +(lastRow/6).toFixed(2);
  Logger.log("Total Unique Users "+totalUniqueUsers)
  activeTab.getRange(1,1,1,1).activate();
  
  /**For loop to transposed the data */
  for (var i = 1; i <= totalUniqueUsers;i++){
    //OFFSET(starting cell, number of rows to shift, no of columns to shift,range of rows)
    activeTab.getCurrentCell().offset(0,0,6,1).copyTo(activeTab.getActiveRange(),SpreadsheetApp.CopyPasteType.PASTE_NORMAL,true);
    activeTab.getCurrentCell().offset(0, 0, 7, 1).activate();
    activeTab.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    activeTab.getCurrentCell().offset(6, 0).activate();
  }
  
  /**Create a Column heading */
  activeTab.insertRowsBefore(activeTab.getRange(1,1).getRow(),1);
  activeTab.getRange("B1").setValue('Copy Data');

  var lastColumn = activeTab.getLastColumn();
  var maxRows = activeTab.getMaxRows();

  /**Create a filter to sort by values first and blanks last */
  activeTab.getRange(1,1,maxRows,lastColumn).createFilter();
  activeTab.getRange('B1').activate();
  activeTab.getFilter().sort(2, true);
  activeTab.getFilter().remove();

/**Add additional Columns */
  activeTab.getRange("A1").setValue('GES Active?');
  activeTab.getRange("B1").setValue('Fullname');
  activeTab.getRange("C1").setValue('Email');
  activeTab.getRange("D1").setValue('Progress');
  activeTab.getRange("E1").setValue('Date Enrolled');
  activeTab.getRange("F1").setValue('Status');

  activeTab.insertColumnsBefore(activeTab.getRange('C:C').getColumn(),2);
  activeTab.insertColumnsBefore(activeTab.getRange('F:F').getColumn(),1);
  activeTab.getRange("C1").setValue('Firstname');
  activeTab.getRange("D1").setValue('Lastname');
  activeTab.getRange("F1").setValue('Valid Email ?');

  lastRow = activeTab.getLastRow();
/**Add formulas in cells to check for firstname and lastname and place in correct format */
  activeTab.getRange('C2').setFormula('=PROPER(LEFT(B2,SEARCH(" ",B2)-1)) ');
  activeTab.getRange('C2').copyTo(activeTab.getRange(2,3,lastRow-1));

  activeTab.getRange('D2').setFormula('=PROPER(RIGHT(B2,LEN(B2)-SEARCH(" ",B2)))');
  activeTab.getRange("D2").copyTo(activeTab.getRange(2,4,lastRow-1));

  activeTab.getRange('F2').setFormula('=ISEMAIL(E2)');
  activeTab.getRange("F2").copyTo(activeTab.getRange(2,6,lastRow-1))

  /**add formulas to begin the process of cleaning invalid emails, first by searching for valid emails from truncated emails spreadsheet */
  activeTab.insertColumnsBefore(activeTab.getRange('F:F').getColumn(),3);
  activeTab.getRange('H2').setFormula('=IFNA(VLOOKUP(E2,\'Truncated Emails\'!A:B,2,FALSE),"Valid email")');
  activeTab.getRange('H2').copyTo(activeTab.getRange(2,8,lastRow-1));

/**Copy and paste the correct emails from truncated spreadsheet to the corrsesponding row */
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).createFilter();
  var filter_criteria_truncated = SpreadsheetApp.newFilterCriteria().whenTextDoesNotContain('Valid Email');
  activeTab.getFilter().setColumnFilterCriteria(8,filter_criteria_truncated);
  activeTab.getRange('H:H').copyTo(activeTab.getRange('E:E').activate(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  activeTab.getFilter().remove();
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).createFilter();

 /**Cleaning additional truncated emails that contain @g and changing it to @gmail.com */
  var filter_criteria_valid = SpreadsheetApp.newFilterCriteria().whenTextContains('FALSE');
  activeTab.getFilter().setColumnFilterCriteria(9,filter_criteria_valid);
  var filter_criteria_gmail = SpreadsheetApp.newFilterCriteria().whenTextContains('@g')
  activeTab.getFilter().setColumnFilterCriteria(5,filter_criteria_gmail);
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).activate();
  activeTab.getRange('E:E').activate();
  activeTab.getRange('E:E').splitTextToColumns('@');
  activeTab.getRange('F:F').activate();
  activeTab.getCurrentCell().setValue('@gmail.com');
  activeTab.getRange('F1').activate();
  activeTab.getRange('F1').copyTo(activeTab.getRange('F:F'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  activeTab.getRange('G:G').activate();
  activeTab.getCurrentCell().setFormula('=CONCATENATE(E1,F1)');
  
  var lr = activeTab.getLastRow();
  activeTab.getRange('G1').copyTo(activeTab.getRange(2,7,lr-1));
  
  var currentCell = activeTab.getCurrentCell();
  activeTab.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  activeTab.getRange(1,7,lr-1).copyTo(activeTab.getRange('E:E'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  var app = SpreadsheetApp;
  app.getActiveSpreadsheet().getActiveSheet().getFilter().remove();

  /**Cleaning additional truncated emails that contain @y and changing it to @yahoo.com */
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).createFilter();
  var filter_criteria_valid = SpreadsheetApp.newFilterCriteria().whenTextContains('FALSE');
  activeTab.getFilter().setColumnFilterCriteria(9,filter_criteria_valid);

  var filter_criteria_yahoo = SpreadsheetApp.newFilterCriteria().whenTextContains('@y')
  activeTab.getFilter().setColumnFilterCriteria(5,filter_criteria_yahoo);
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).activate();
  
  activeTab.getRange('E:E').activate();
  activeTab.getRange('E:E').splitTextToColumns('@');
  activeTab.getRange('F:F').activate();
  activeTab.getCurrentCell().setValue('@yahoo.com');

  activeTab.getRange('F1').activate();
  activeTab.getRange('F1').copyTo(activeTab.getRange(2,6,activeTab.getLastRow()), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  activeTab.getRange('G:G').activate();
  activeTab.getCurrentCell().setFormula('=CONCATENATE(E1,F1)');
  activeTab.getRange('G1').copyTo(activeTab.getRange(2,7,activeTab.getLastRow()-1));
  var currentCell = activeTab.getCurrentCell();
  activeTab.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  
  activeTab.getRange(1,7,activeTab.getLastRow()-1).copyTo(activeTab.getRange('E:E'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);


  var app = SpreadsheetApp;
  app.getActiveSpreadsheet().getActiveSheet().getFilter().remove();
  activeTab.deleteColumns(6,3);
  activeTab.getRange("E1").setValue('Email');

/**Cleanup by removing whitespaces and duplicates*/
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).createFilter();
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).trimWhitespace();
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).offset(1, 0, activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).getNumRows() - 1).activate();
  activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).removeDuplicates().activate();
  activeTab.deleteColumns(1);
  app.getActiveSpreadsheet().getActiveSheet().getFilter().remove();
  //activeTab.getRange(1,1,activeTab.getLastRow(),activeTab.getLastColumn()).createFilter();
  

};
