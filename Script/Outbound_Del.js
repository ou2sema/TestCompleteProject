var Login = require("Login");
var Check_position = require("Check_position");

let my_saved_file =Excel.Open("C:\\test_data\\order_saved.xlsx");
function main()
  {
    
     switch ( Check_position.check_wind() )
  {
    case 1:
    outband_trans();
      break;
    case 2:
      Login.main();
      outband_trans();
      break;
   case 3:
        TestedApps.saplogon.Run();
        Login.main();
        outband_trans();
      break;
  }
     

  
function  outband_trans()
{
  aqTestCase.Begin("Outband Delivery Number");
  try{
       
  Aliases.saplogon.mainwindowSap.toolbarTbar0.okcodefieldOkcd.Keys("vl10e[Enter]");
  Aliases.saplogon.mainwindowSap.userareaUsr.CTextField("ctxtP_LERUL").Click()
  Aliases.saplogon.mainwindowSap.userareaUsr.CTextField("ctxtP_LERUL").Keys("[BS]");
  Aliases.saplogon.mainwindowSap.userareaUsr.TabStrip("tabsTABSTRIP_ORDER_CRITERIA").Tab("tabpS0S_TAB2").Click();
  Aliases.saplogon.mainwindowSap.userareaUsr.TabStrip("tabsTABSTRIP_ORDER_CRITERIA").Tab("tabpS0S_TAB2").ScrollContainer("ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1020").CTextField("ctxtST_VBELN-LOW").Keys("600002842");
  Aliases.saplogon.mainwindowSap.Toolbar("tbar[1]").Button("btn[8]").Click();
  Test1();
 /* Aliases.saplogon.mainwindowSap.userareaUsr.CheckBox("chk[1,5]").Click();
  Aliases.saplogon.mainwindowSap.Toolbar("tbar[1]").Button("btn[20]").Click();
  Aliases.saplogon.mainwindowSap.userareaUsr.CheckBox("chk[1,6]").Click();
  Aliases.saplogon.mainwindowSap.Toolbar("tbar[1]").Button("btn[18]").Click();
  Aliases.saplogon.mainwindowSap.userareaUsr.Label("lbl[42,6]").Click();*/
  let outband_del_bnr =Aliases.saplogon.mainwindowSap.userareaUsr.SimpleContainer("subSUBSCREEN_HEADER:SAPMV50A:1502").CTextField("ctxtLIKP-VBELN").Text;
  let outband_sh=my_saved_file.SheetByTitle("outband");
  let rowIndex = outband_sh.RowCount + 1;
  outband_sh.Cell("A", rowIndex).Value = outband_del_bnr;
   my_saved_file.Save();
   Aliases.saplogon.mainwindowSap.toolbarTbar0.Button("btn[12]").Click();
    Aliases.saplogon.mainwindowSap.toolbarTbar0.Button("btn[12]").Click();
     Aliases.saplogon.mainwindowSap.toolbarTbar0.Button("btn[12]").Click();
 


module.exports.outband_trans = outband_trans;
 }
    catch (e){
   
     }

  finally{
    aqTestCase.End();
        }
}
}
module.exports.main = main;