
  aqTestCase.Begin("successful login");
    try
    {
    function single_sign_on(item)
      {
  /* let listView = Aliases.saplogon.dlgSAPLogon760.AfxOleControl140.AfxMDIFrame140.AfxFrameOrView140.myContainer.ListView;
   listView.ClickItemR(item, 0);
  listView.PopupMenu.Click("SNC Logon Without Single Sign-On ");*/
  //Test1()
  let listView = Aliases.saplogon.dlgSAPLogon770.AfxOleControl140.AfxMDIFrame140.AfxFrameOrView140.myContainer.ListView;
  listView.ClickItemR(item, 0);
  listView.PopupMenu.Click("SNC Logon Without Single Sign-On");
      }
function login(clnt,usr,pwd)
{
  
    let saplogon = Aliases.saplogon;
  let userarea = saplogon.mainwindowSap.userareaUsr;
  let textfield = userarea.textfieldClient;
  textfield = userarea.textfieldUser;
  Aliases.saplogon.mainwindowSap.userareaUsr.textfieldClient.Keys(clnt);
  userarea.textfieldUser.Keys(usr);
  userarea.passwordfieldPassword.Keys(pwd);
   
   
  if(textfield.MaxLength==0 && userarea.textfieldUser.MaxLength==0 && userarea.passwordfieldPassword.MaxLength==0){
    MessageDlg("please check your input data !",mtConfirmation,mbOK,0);
    Runner.Stop();
    
 }
 userarea.passwordfieldPassword.Keys( "[Enter]");
  
} 
module.exports.single_sign_on = single_sign_on;
module.exports.login = login;

 }
    catch (e){
   
     }

  finally{
    // Sets the end of the "Create order" test
  
aqTestCase.End();
        }

