
var Create_Order = require("Create_Order");  
function main()
{
  var fileName = "C:\\test_data\\data1.xlsx";
  
  // Open the Excel file
  var excelFile = Excel.Open(fileName);
  
  // Search the specified value in the Excel file
  //for (var i = 0; i < excelFile.SheetCount; i++)
  //{
    var excelSheet = excelFile.SheetByTitle("trans");
    for (var j = 1; j <= excelSheet.ColumnCount; j++)
    {
      for (var k =1; k <= excelSheet.RowCount; k++)
      {
      switch ( excelSheet.Cell(j, k).Value )
      {
        case "VA01" :
            Log.Message(excelSheet.Cell(j, k).Value);
            Create_Order.main();
          break;
        case "test" :
            Log.Message(excelSheet.Cell(j, k).Value);
          break;
          
      }
     
      }
      
    }
     
    Log.SaveResultsAs("C:\\test_data\\log\\report.zip", lsPackedHTML);
}

