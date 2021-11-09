/*function ReadDataFromExcel()
{
  let Excel1 = getActiveXObject("Excel.Application");
  Excel1.Workbooks.Open("C:\\test_data\\data.xlsx");
  Excel1.Sheets[1].ActiveSheet;
  let RowCount = Excel1.ActiveSheet.UsedRange.Rows.Count;
  let ColumnCount = Excel1.ActiveSheet.UsedRange.Columns.Count;
  
  for (let i = 1; i <= RowCount; i++)
  {
    let s = "";
    for (let j = 1; j <= ColumnCount; j++)
      s += (VarToString(Excel1.Cells.Item(i, j)) + "\r\n");
    Log.Message("Row: " + i, s);
  }

  Excel1.Quit();<
}
*/

//test function
function data_file()
{
  var fileName = "C:\\test_data\\data1.xlsx";
  var searchValue = "VA01";
  
  // Open the Excel file
  var excelFile = Excel.Open(fileName);
  
  // Search the specified value in the Excel file
  var found = false;
  //for (var i = 0; i < excelFile.SheetCount; i++)
  //{
    var excelSheet = excelFile.SheetByTitle("trans");
    for (var j = 1; j <= excelSheet.ColumnCount; j++)
    {
      for (var k =1; k <= excelSheet.RowCount; k++)
      if (excelSheet.Cell(j, k).Value == searchValue)
      {
        Log.Message("The specified value is found on the " + excelSheet.Title + " sheet in the (" + j + "," + k + ") cell");
         Log.Message(excelSheet.Cell(j, k).Value);
        found = true;
      }
    }
  //}
  if (found == false)
    Log.Message("The specified value isn't found");
}