using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BudgetExcelSheets
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            SaltireAPI.Global.DataConnectionString = "Server=Vision-fp1; Database=Saltire_Vision; User Id=crystalreports; Password=crystalreports";
            SaltireAPI.Global.sqlConnection = new System.Data.SqlClient.SqlConnection(SaltireAPI.Global.DataConnectionString);
            SaltireAPI.Global.DateFormatString = "dd/MMM/yyyy HH:mm:ss.fff";

            Application.Run(new frmMain());
        }
    }
}
