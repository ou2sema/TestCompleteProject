let x=0;
let y=0;
let z=0;
function check_wind()
{
  

  if (!(Aliases.saplogon.dlgSAPLogon770.Exists))
  {
     x=3;
     //TestedApps.saplogon.Run();
     Log.Message("SAP is not running ,verif= "+x);
     return x;
  }
  
   if ((Aliases.saplogon.dlgSAPLogon770.Exists && !(Aliases.saplogon.mainwindowSap.Exists) ))
  {
    
  y=2;
   Log.Message("SAP login page ,verif= "+y);
  return y;
  }
  if(Aliases.saplogon.mainwindowSap.Exists)
  {
    //Start_SAP.start_sap();
    z=1;
    Log.Message("TA1 home page ,verif= "+z);
   return z;
  }
  
 
}
module.exports.check_wind = check_wind;