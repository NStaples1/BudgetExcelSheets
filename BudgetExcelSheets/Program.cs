using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using static DevExpress.XtraPrinting.Native.ExportOptionsPropertiesNames;

namespace BudgetExcelSheets
{
   internal static class Program
   {
      /// <summary>
      /// The main entry point for the application.
      /// </summary>
      [STAThread]
      static void Main(string[] args)
      {
         Application.EnableVisualStyles();
         Application.SetCompatibleTextRenderingDefault(false);
         SaltireAPI.Global.DataConnectionString = "Server=Vision-fp1; Database=Saltire_Vision; User Id=crystalreports; Password=crystalreports";
         SaltireAPI.Global.sqlConnection = new System.Data.SqlClient.SqlConnection(SaltireAPI.Global.DataConnectionString);
         SaltireAPI.Global.DateFormatString = "dd/MMM/yyyy HH:mm:ss.fff";

         if (args.Length > 0)
         {
            try
            {
               Classes.Global.hasArgs = true;
               new frmMain().CreateSpreadsheet();
            }
            catch (Exception ex)
            {
               frmMain.Create_Error_Log("Program", "Main", ex.Message, ex);
            }
         }
         else
         {
            Classes.Global.hasArgs = false;
            string Filename = Application.StartupPath + @"\Application Version Number.json";
            bool ExitApplication = false;

            if (File.Exists(Filename) && System.Environment.MachineName != "VISION-W10-DT46" && System.Environment.MachineName != "VISION-W10-DT52")
            {
               VPS.Core.VersionChecking vCheck = new VPS.Core.VersionChecking();
               VPS.Core.VersionNumber vNumber = Newtonsoft.Json.JsonConvert.DeserializeObject<VPS.Core.VersionNumber>(File.ReadAllText(Filename));
               if (vNumber != null)
               {
                  if (vCheck.RequiresUpdate(vNumber.Version_Number, @"\\Vision-fp1\CompanyData\IT\Development\Updates\Monthly Sales Analysis"))
                  {
                     /************************************************************************************************************************************
                      * Load the application Updater
                      ***********************************************************************************************************************************/
                     string Params = "\"" + Application.ExecutablePath + "\" \"" + vNumber.Version_Number + "\" \"\\\\Vision-fp1\\CompanyData\\IT\\Development\\Updates\\Monthly Sales Analysis\"" + " \"" + Path.GetDirectoryName(Application.ExecutablePath) + "\"";
                     Process.Start(@"\\Vision-fp1\companydata\IT\Development\Updates\Update Application\Automatic Updater.exe", Params);
                     ExitApplication = true;
                  }
               }
            }

            if (ExitApplication)
               Application.Exit();
            else
            {
               Application.Run(new frmMain());
            }
         }
      }
   }
}
