using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace DXTools
{
	/// <summary>
	/// Summary description for ProcessError.
	/// </summary>
	public class ProcessError
	{
      private const string ProgramName = "DXTools";

		public ProcessError()
		{
         
		}

      public static void Show(Form myForm, string strProcedureName, Exception e, List<string> AdditionalData = null)
      {
         string ErrorMessage;
         ErrorMessage = "An error has occurred in the application. \n" +
            "Module: " + myForm.Name + " \n" +
            "Procedure: " + strProcedureName + " \n" +
            "Line Number: " + linenumber(e) + " \n" +
            "Error Message: \n\n" + e.Message + "\n\n" +
            "Please contact your system's administrator. \n";
         MessageBox.Show(myForm, ErrorMessage, ProgramName + " Application Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
      }

      public static void Show(UserControl myControl,string strProcedureName,Exception e, List<string> AdditionalData = null)
		{
         string ErrorMessage;
			ErrorMessage = "An error has occurred in the application. \n" + 
				"Module: " + myControl.Name + " \n" +
				"Procedure: " + strProcedureName + " \n" +
            "Line Number: " + linenumber(e) + " \n" +
            "Error Message: \n\n" + e.Message + "\n\n" +
				"Please contact your system's administrator. \n";
         MessageBox.Show(myControl, ErrorMessage, ProgramName + " Application Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
      }

      public static void Show(string myClass,string strProcedureName,Exception e, List<string> AdditionalData = null)
		{
         string ErrorMessage;
			ErrorMessage = "An error has occurred in the application. \n" + 
				"Module: " + myClass + " \n" +
				"Procedure: " + strProcedureName + " \n" +
            "Line Number: " + linenumber(e) + " \n" +
            "Error Message: \n\n" + e.Message + "\n\n" +
				"Please contact your system's administrator. \n";
         ShowMessageBox(ErrorMessage);
      }

      private static void ShowMessageBox(string Message)
      {
         var thread = new System.Threading.Thread(() => { MessageBox.Show(Message, ProgramName + " Application Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); });
         thread.Start();
         while (thread.IsAlive)
         {
            // Do nothing until messagebox is closed
            System.Threading.Thread.Sleep(200);
         }
      } 

      public static void Show(UserControl Owner, string myClass, string strProcedureName, Exception e, List<string> AdditionalData = null)
      {
         string ErrorMessage;
         ErrorMessage = "An error has occurred in the application. \n" +
            "Module: " + myClass + " \n" +
            "Procedure: " + strProcedureName + " \n" +
            "Line Number: " + linenumber(e) + " \n" +
            "Error Message: \n\n" + e.Message + "\n\n" +
            "Please contact your system's administrator. \n";
         MessageBox.Show(Owner, ErrorMessage, ProgramName + " Application Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
      }

      public static void Show(Form Owner, string myClass, string strProcedureName, Exception e, List<string> AdditionalData = null)
      {
         string ErrorMessage;
         ErrorMessage = "An error has occurred in the application. \n" +
            "Module: " + myClass + " \n" +
            "Procedure: " + strProcedureName + " \n" +
            "Line Number: " + linenumber(e) + " \n" +
            "Error Message: \n\n" + e.Message + "\n\n" +
            "Please contact your system's administrator. \n";
         MessageBox.Show(Owner, ErrorMessage, ProgramName + " Application Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
      }

      private static string linenumber(Exception ex)
      {
         int iIndex = ex.StackTrace.LastIndexOf("line ");
         string linenumber;
         if (iIndex >= 0)
            linenumber = ex.StackTrace.Substring(iIndex).Replace("line", "");
         else
            linenumber = "Unknown";
         linenumber = linenumber.Trim().ToString();
         return linenumber;
      }

		public static void Message(Form myForm,string myMessage)
		{
         MessageBox.Show(myForm, myMessage, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Question);
		}

		public static void Message(string myMessage)
		{
         MessageBox.Show(myMessage, ProgramName, MessageBoxButtons.OK, MessageBoxIcon.Question);
		}

      public static string Return_Error(Form myForm, string strProcedureName, Exception e, List<string> AdditionalData = null)
      {
         return Return_Error(myForm.Name, strProcedureName, e, AdditionalData);
      }

      public static string Return_Error(string myClass, string strProcedureName, Exception e, List<string> AdditionalData = null)
      {
         string ErrorMessage;
         ErrorMessage = "An error has occurred in the application. \n" +
            "Module: " + myClass + " \n" +
            "Procedure: " + strProcedureName + " \n" +
            "Line Number: " + linenumber(e) + " \n" +
            "Error Message: \n\n" + e.Message + "\n\n" +
            "Please contact your system's administrator. \n";

         return ErrorMessage;
      }

      public static string Clean_String(string InputString)
      {
         if (InputString == null || InputString.ToString().Length == 0)
            InputString = null;
         else
         {
            InputString = InputString.Replace("'", "''");
         }
         return InputString;
      }

      private static string GenerateAdditionalData(List<string> AdditionalData)
      {
         if (AdditionalData == null)
            return string.Empty;
         else
         {
            return Clean_String(string.Join(Environment.NewLine + Environment.NewLine + "***********************************************************************" + Environment.NewLine + Environment.NewLine, AdditionalData));
         }
      }
   }
}