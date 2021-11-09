let my_file =Excel.Open("C:\\test_data\\data.xlsx");

function login_attemp1()
{
  let login_sh=my_file.SheetByTitle("login");
  let client=login_sh.Cell("A", 2).Value;
  let user=login_sh.Cell("B", 2).Value;
  let pwd=login_sh.Cell("C", 2).Value;
  let listView = Aliases.saplogon.dlgSAPLogon760.AfxOleControl140.AfxMDIFrame140.AfxFrameOrView140.myContainer.ListView;
  listView.ClickItemR("306 - TA1: Sandbox synapsis: S/4HANA", 0);
  listView.PopupMenu.Click("SNC Logon Without Single Sign-On ");
  let saplogon = Aliases.saplogon;
  let userarea = saplogon.mainwindowSap.userareaUsr;
  let textfield = userarea.textfieldClient;
  textfield = userarea.textfieldUser;
  Aliases.saplogon.mainwindowSap.userareaUsr.textfieldClient.Keys(client);
  userarea.textfieldUser.Keys(user);
  userarea.passwordfieldPassword.Keys(pwd+"[Enter]");
  Aliases.saplogon.mainwindowSap.Close();
  Aliases.saplogon.modalwindowSalesDocumentType143E.userareaUsr.buttonYes.Click();
}
function login_attemp2()
{
  let wrong =my_file.SheetByTitle("wrong_data");
  let clnt=wrong.Cell("A", 2).Value;
  let usr=wrong.Cell("B", 2).Value;
  let pwd=wrong.Cell("C", 2).Value;
  let listView = Aliases.saplogon.dlgSAPLogon760.AfxOleControl140.AfxMDIFrame140.AfxFrameOrView140.myContainer.ListView;
  listView.ClickItemR("306 - TA1: Sandbox synapsis: S/4HANA", 0);
  listView.PopupMenu.Click("SNC Logon Without Single Sign-On ");
  let saplogon = Aliases.saplogon;
  let userarea = saplogon.mainwindowSap.userareaUsr;
  let textfield = userarea.textfieldClient;
  textfield = userarea.textfieldUser;
  Aliases.saplogon.mainwindowSap.userareaUsr.textfieldClient.Keys(clnt);
  userarea.textfieldUser.Keys(usr);
  userarea.passwordfieldPassword.Keys(pwd+"[Enter]");
  Aliases.saplogon.mainwindowSap.Close();
  Aliases.saplogon.modalwindowSalesDocumentType143E.userareaUsr.buttonYes.Click();
   
 }
module.exports.login_attemp1 = login_attemp1;
module.exports.login_attemp2 = login_attemp2;