var main_menu = require("main_menu");
let my_file =Excel.Open("C:\\test_data\\data.xlsx");
let my_saved_file =Excel.Open("C:\\test_data\\order_saved.xlsx");
let login_sh=my_file.SheetByTitle("login");
let client=login_sh.Cell("A", 3).Value;
let user=login_sh.Cell("B", 3).Value;
let pwd=login_sh.Cell("C", 3).Value;
function main()
{
   main_menu.single_sign_on("308 - TA5: Sandbox synapsis: GTS");
   main_menu.login()
}
module.exports.main = main;

/*function Test1()
{
  let listView = Aliases.saplogon.dlgSAPLogon760.AfxOleControl140.AfxMDIFrame140.AfxFrameOrView140.myContainer.ListView;
  listView.MouseWheel(-4);
  listView.ClickItemR("308 - TA5: Sandbox synapsis: GTS", 0);
  listView.PopupMenu.Click("SNC Logon Without Single Sign-On ");
}*/