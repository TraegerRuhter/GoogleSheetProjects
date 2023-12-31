function ArchiveAndReset() {
  var retail_ready_workbook = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = retail_ready_workbook.getActiveSheet();
  var archiveSheet = retail_ready_workbook.getSheetByName("Archive");


  // get amount of rows to copy except the bottom row
  var totalRows = sourceSheet.getMaxRows();
  var numRowsToCopy = totalRows - 1; // Copy everything except the last row

  // get the default state range
  var defaultStateRange = sourceSheet.getRange(1, 5, numRowsToCopy, 2); 
  
  /*
  Define the range to copy (excluding the last row)
  - Start Row (1st parameter: 1): The row index where the range begins. 
    '1' indicates that the range starts from the first row of the sheet.
  - Start Column (2nd parameter: 3): The column index where the range begins. 
    Columns are indexed starting from 1, so '3' refers to the third column of the sheet (column 'C').
  - Number of Rows (3rd parameter: numRowsToCopy): Specifies how many rows will be included in the range. 
    'numRowsToCopy' is calculated as the total number of rows in the sheet minus one (to exclude the last row).
  - Number of Columns (4th parameter: 2): Indicates how many columns the range will span. 
    '2' means that the range will include two columns (C and D).
  */
  var sourceRange = sourceSheet.getRange(1, 3, numRowsToCopy, 2); 

  // Insert 2 new columns at B and C in the Archive sheet
  archiveSheet.insertColumnsBefore(2, 2);

  // Copy the range to the Archive sheet, starting at the first row and second column (B)
  sourceRange.copyTo(archiveSheet.getRange(1, 2, numRowsToCopy, 2), {contentsOnly: false});

  // set the copyed colums to the default
  defaultStateRange.copyTo(sourceSheet.getRange(1, 3, numRowsToCopy, 2), {contentsOnly: false});

}
