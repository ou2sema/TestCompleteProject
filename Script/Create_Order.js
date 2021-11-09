var Login = require("Login");
var Check_position = require("Check_position");
//Static and global Variables
var main_menu = require("main_menu");
let my_file =Excel.Open("C:\\test_data\\data1.xlsx");
let my_saved_file =Excel.Open("C:\\test_data\\order_saved.xlsx");
let login_sh=my_file.SheetByTitle("login");
let sys_sh=my_file.SheetByTitle("trans");
let client;
let user;
let pwd;
let ta_sys;
// Sets the end of the "Create order" test
function main()
  {
    let arr=data();
  switch ( Check_position.check_wind() )
  {
    case 1:
     main_menu.login(arr[1],arr[2],arr[3]);
    OrderCreation();
      break;
    case 2:
    Log.Message(arr[0]);
      main_menu.single_sign_on(arr[0]);
      main_menu.login(arr[1],arr[2],arr[3]);
      OrderCreation();
      break;
   case 3:
        TestedApps.saplogon.Run(1, true);
        //Sys.Process("saplogon")
        main()
      break;
  }
  

function data()
{
  
  //let login_sh=my_file.SheetByTitle("login");
  client=login_sh.Cell("A", 2).Value;
  user=login_sh.Cell("B", 2).Value;
  pwd=login_sh.Cell("C", 2).Value;
  ta_sys=sys_sh.Cell("B", 2).Value;
  //ta_sys="306 - TA1: Sandbox synapsis: S/4HANA"
  return [ta_sys,client,user,pwd];
}
  function OrderCreation()
{
  aqTestCase.Begin("Create an order");
    try
 {
  let sales_doc_sh=my_file.SheetByTitle("sales_doc");
  let order_type=sales_doc_sh.Cell("A", 2).Value;
  let sales_org=sales_doc_sh.Cell("B", 2).Value;
  let dist_ch=sales_doc_sh.Cell("C", 2).Value;
  let div=sales_doc_sh.Cell("D", 2).Value;
  let sales_off=sales_doc_sh.Cell("E", 2).Value;
  
  
  let okcodefield = Aliases.saplogon.mainwindowSap.toolbarTbar0.okcodefieldOkcd;
  okcodefield.Click(37, 7, skCtrl);
  okcodefield.Click(64, 13);
  okcodefield.Keys("va01[Enter]");
  let userarea1 = Aliases.saplogon.mainwindowSap.userareaUsr;
  let ctextfield = userarea1.textfieldOrderType;
  //ctextfield.Click(25, 2);
  
  ctextfield.Keys("[Del][Del][Del][Del][Del]"+order_type+"[Tab]");
  ctextfield = userarea1.textfieldSalesOrganization;
  ctextfield.Keys("[BS]"+sales_org+"[Tab]");
  userarea1.textfieldDistributionChannel.Keys("[BS]"+dist_ch+"[Tab]");
  userarea1.textfieldDivision.Keys("[BS]"+div+"[Tab]");
  let ctextfield2 = userarea1.textfieldSalesOffice;
  ctextfield2.Keys("[BS]"+sales_off+"[Enter]");
  FillOrderPart1();
}  
catch (e){
   
     }

  finally{
    aqTestCase.End();
        }
  }
      
    
   function FillOrderPart1()
{
  let std_order_sh=my_file.SheetByTitle("std_order");
  let sold_part= std_order_sh.Cell("A", 2).Value;
  let ship_part= std_order_sh.Cell("B", 2).Value;
  let cust_ref= std_order_sh.Cell("C", 2).Value;
  let cust_ref_date= std_order_sh.Cell("D", 2).Value;
  


    
  
  let saplogon = Aliases.saplogon;
  let simplecontainer = saplogon.mainwindowSap.userareaUsr.simplecontainerSubscreenHeaderSa;
  let simplecontainer2 = simplecontainer.simplecontainerPartSubSapmv45a47;
  let ctextfield1 = simplecontainer2.textfieldSoldToParty;
  ctextfield1.Click(33, 10);
  ctextfield1.Keys(sold_part);
  ctextfield1 = simplecontainer2.textfieldShipToParty;
  ctextfield1.Click(33, 9);
  ctextfield1.Keys(ship_part);
   Aliases.saplogon.mainwindowSap.userareaUsr.simplecontainerSubscreenHeaderSa.textfieldCustRefDate.Keys(cust_ref_date)
  let textfield1 = simplecontainer.textfieldCustReference;
  //textfield1.Click(33, 14);
  textfield1.Keys(cust_ref+"[Enter]");
  Aliases.saplogon.modalwindowSalesDocumentType143E.toolbarTbar0.buttonBtn0.Keys("[Enter]")
  FillOrderPart2();
  }
   function FillOrderPart2()
{
  
  //item data sheet
  let item_data_sh=my_file.SheetByTitle("item_data");
  let order_qty=item_data_sh.Cell("A", 2).Value;
  let del_date=item_data_sh.Cell("B", 2).Value;
  
  
  //std data sheet
  let std_order_sh=my_file.SheetByTitle("std_order");
  let material=std_order_sh.Cell("E", 2).Value;
  let cust_mat_nb=std_order_sh.Cell("F", 2).Value;
  let item_nb=std_order_sh.Cell("G", 2).Value;
//req deliv date
  Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripOverview.tabT01.scrollcontainerSubscreenBodySapm.ScrollContainer("ssubHEADER_FRAME:SAPMV45A:4440").CTextField("ctxtRV45A-KETDAT").Keys("[Enter]")
  let saplogon = Aliases.saplogon;
  let userarea = saplogon.mainwindowSap.userareaUsr;
  userarea.simplecontainerSubscreenHeaderSa.textfieldCustRefDate.Keys("[Enter]");
  let modalwindow = saplogon.modalwindowSalesDocumentType143E;
  let tablecontrol = userarea.tabstripTaxiTabstripOverview.tabT01.scrollcontainerSubscreenBodySapm.simplecontainerSubscreenTcSapmv4.tablecontrolSapmv45atctrlUErfAuf;
  let ctextfield = tablecontrol.textfieldRv45aMabnr;
  Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripOverview.tabT01.scrollcontainerSubscreenBodySapm.simplecontainerSubscreenTcSapmv4.tablecontrolSapmv45atctrlUErfAuf.TextField("txtVBAP-POSNR[0,0]").Keys(item_nb+"Tab");
  ctextfield.Keys(material+"[Tab]");
  tablecontrol.textfieldVbapKdmat.Keys(cust_mat_nb+"[Enter]");
  ctextfield.DblClick(105, 7);
  let tab = userarea.tabstripTaxiTabstripItem.tabT07;
  tab.Click(64, 10);
  tablecontrol = tab.scrollcontainerSubscreenBodySapm.tablecontrolSapmv45atctrlPein;
  let textfield = tablecontrol.textfieldVbepWmeng;
  Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripItem.tabT07.scrollcontainerSubscreenBodySapm.tablecontrolSapmv45atctrlPein.textfieldVbepWmeng.Keys(order_qty)
  Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripItem.tabT07.scrollcontainerSubscreenBodySapm.tablecontrolSapmv45atctrlPein.textfieldRv45aEtdat.Keys(del_date);
  FillOrderPart3();
}


   function FillOrderPart3()
{
  let item_data_sh=my_file.SheetByTitle("item_data");
   let rout=item_data_sh.Cell("C", 2).Value;
  let cn_ty=item_data_sh.Cell("D", 2).Value;
  let amount=item_data_sh.Cell("E", 2).Value;
  let per=item_data_sh.Cell("F", 2).Value;
  let saplogon = Aliases.saplogon;
  let mainwindow = saplogon.mainwindowSap;
  let userarea = mainwindow.userareaUsr;
  let tabstrip = userarea.tabstripTaxiTabstripItem;
  tabstrip.tabT07.scrollcontainerSubscreenBodySapm.tablecontrolSapmv45atctrlPein.textfieldRv45aEtdat.Keys("[Enter]");
  let tab = tabstrip.tabT03;
  tab.Click(36, 11);
  let ctextfield = tab.scrollcontainerSubscreenBodySapm.textfieldRoute;
  ctextfield.Click(30, 9);
  ctextfield.Keys(rout+"[Enter]");
  tab = tabstrip.tabT05;
  tab.Click(21, 4);
  tab.Click(26, 7);
  let tablecontrol = tab.scrollcontainerSubscreenBodySapl.tablecontrolSaplv69atctrlKonditi;
  //tablecontrol.textfieldKomvKschl.SetText(cn_ty+"[Tab]");
  Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripItem.tabT05.scrollcontainerSubscreenBodySapl.tablecontrolSaplv69atctrlKonditi.textfieldKomvKschl2.Keys(cn_ty+"[Tab]");
  
  
  //tablecontrol.textfieldKomvKbetr.Keys(amount+"[Tab]");
  Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripItem.tabT05.scrollcontainerSubscreenBodySapl.tablecontrolSaplv69atctrlKonditi.textfieldKomvKbetr2.Keys(amount+"[Tab]");
  //tablecontrol.textfieldRv61aKoein.Keys("[Tab]");
 // tablecontrol.textfieldKomvKpein.Keys(per+"[Enter]");
  Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripItem.tabT05.scrollcontainerSubscreenBodySapl.tablecontrolSaplv69atctrlKonditi.textfieldKomvKpein2.Keys(per+"[Enter]");
  //new item PR01
  PR01()
  tabstrip.tabT12.Click(0, 12);
  tabstrip.tabT13.Click(117, 12);
  tabstrip.tabT14.Click(12, 7);
  tab = tabstrip.tabT15;
  tab.Click(94, 6);
  let simplecontainer = tab.scrollcontainerSubscreenBodySapm.simplecontainerCustomerScreenSap;
  simplecontainer.comboboxCarlineDrXlmaier.Click(175, 10);
  let listBox = saplogon.Afx.ListBox;
  listBox.ClickItemXY(0, 69, 7);
  simplecontainer.comboboxProductType.Click(168, 11);
  listBox.ClickItemXY(0, 52, 5);
  
  Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripItem.tabT15.scrollcontainerSubscreenBodySapm.simplecontainerCustomerScreenSap.comboboxProductType.Keys("^[F8]");
    Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripItem.tabT15.scrollcontainerSubscreenBodySapm.simplecontainerCustomerScreenSap.comboboxCarlineDrXlmaier.Keys("^s");
  OrderSaved();
}


function OrderSaved()
{
  let order_sh=my_saved_file.SheetByTitle("order");
  let rowIndex = order_sh.RowCount + 1;
 let result=Aliases.saplogon.mainwindowSap.Statusbar("sbar").Text;
 let str =result.substr(34,40);
 let res=str.localeCompare("saved.");
   if (res==0)
   {
   order_sh.Cell("A", rowIndex).Value = result.substr(15,9);
   my_saved_file.Save();
 Aliases.saplogon.mainwindowSap.toolbarTbar0.Button("btn[12]").Click();
 Aliases.saplogon.mainwindowSap.toolbarTbar0.Button("btn[12]").Click();
 Aliases.saplogon.mainwindowSap.Close();
 Aliases.saplogon.modalwindowSalesDocumentType143E.userareaUsr.buttonYes.Click();
  
   //Runner.Stop();
   }
}






//module.exports.init = init;module.exports.single_sign_on = single_sign_on;

module.exports.OrderCreation = OrderCreation;
//module.exports.main = main;

}
module.exports.main = main;

function PR01()
{
  let tablecontrol = Aliases.saplogon.mainwindowSap.userareaUsr.tabstripTaxiTabstripItem.tabT05.scrollcontainerSubscreenBodySapl.tablecontrolSaplv69atctrlKonditi;
  let ctextfield = tablecontrol.textfieldKomvKschl3;
  ctextfield.Click(13, 12);
  ctextfield.Keys("PR01[Tab]");
  let textfield = tablecontrol.textfieldKomvKbetr3;
  textfield.Keys("5,00[Tab]");
  //textfield.Click(22, 14);
  textfield = tablecontrol.textfieldKomvKpein3;
  //textfield.Click(15, 15);
  textfield.Keys("100[Enter]");
}