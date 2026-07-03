using DXTools;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BudgetExcelSheets.Classes
{
    public static class Global
    {
      public static bool hasArgs { get; set; }

      private static string str_EmailHost;
      private static string str_EmailHostUser;
      private static string str_EmailHostPassword;
      private static int int_EmailHostPort;
      private static bool bol_EmailHostSSL;
      private static bool bol_EmailHostUsePort;

      public static bool EmailHostUsePort
      {
         get
         {
            return bol_EmailHostUsePort;
         }
         set
         {
            bol_EmailHostUsePort = value;
         }
      }

      public static bool EmailHostSSL
      {
         get
         {
            return bol_EmailHostSSL;
         }
         set
         {
            bol_EmailHostSSL = value;
         }
      }

      public static int EmailHostPort
      {
         get
         {
            return int_EmailHostPort;
         }
         set
         {
            int_EmailHostPort = value;
         }
      }

      public static string EmailHost
      {
         get
         {
            return str_EmailHost;
         }
         set
         {
            str_EmailHost = value;
         }
      }

      public static string EmailHostUser
      {
         get
         {
            return str_EmailHostUser;
         }
         set
         {
            str_EmailHostUser = value;
         }
      }

      public static string EmailHostPassword
      {
         get
         {
            return str_EmailHostPassword;
         }
         set
         {
            str_EmailHostPassword = value;
         }
      }

      public static string ReadSignature()
      {
         try
         {
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string signature = string.Empty;
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if (diInfo.Exists)
            {
               FileInfo[] fiSignature = diInfo.GetFiles("*.htm");

               if (fiSignature.Length > 0)
               {
                  StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.Default);
                  signature = sr.ReadToEnd();

                  if (!string.IsNullOrEmpty(signature))
                  {
                     string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty).Replace(" ", "%20");
                     signature = signature.Replace(fileName + "_files/", appDataDir + "/" + fileName.Replace("%20", " ") + "_files/");
                  }
               }
            }
            return signature;
         }
         catch (Exception ex)
         {
            // We dont want to show an error but we will log it
            ProcessError.Return_Error("Global", "ReadSignature", ex);
            return string.Empty;
         }
      }

      public static string ConvertToString(object sValue)
        {
            try
            {
                if (sValue != null)
                    return sValue.ToString();
                else
                    return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static decimal ConvertToDecimal(object cValue)
        {
            try
            {
                if (cValue != null)
                {
                    string mValue = ConvertToString(cValue);
                    if (!string.IsNullOrEmpty(mValue))
                        return decimal.Parse(mValue);
                    else return 0;
                }
                else
                    return 0;
            }
            catch
            {
                return 0;
            }
        }
        internal static double DivideNum(double DivNumber, double DividedByNumber)
        {
            try
            {
                if (DivNumber != 0 && DividedByNumber > 0)
                    return DivNumber / DividedByNumber;
                else
                    return 0;
            }
            catch (Exception ex)
            {
                ProcessError.Show("Global", "DivideNum", ex, new List<string>() { "DivNumber = " + DivNumber, "DividedByNumber = " + DividedByNumber });
                return 0;
            }
        }

        internal static decimal DivideNum(object DivNumber, object DividedByNumber)
        {
            try
            {
                decimal DivNumberDecimal = ConvertToDecimal(DivNumber);
                decimal DividedByNumberDecimal = ConvertToDecimal(DividedByNumber);

                if (DivNumberDecimal > 0 && DividedByNumberDecimal > 0)
                    return DivNumberDecimal / DividedByNumberDecimal;
                else
                    return 0;
            }
            catch (Exception ex)
            {
                ProcessError.Show("Global", "DivideNum", ex, new List<string>() { "Decimal Return", "DivNumber = " + DivNumber, "DividedByNumber = " + DividedByNumber });
                return 0;
            }
        }

        internal static double DivideNum(double DivNumber, double DividedByNumber, int Decimal_Precision)
        {
            try
            {
                return Math.Round(DivideNum(DivNumber, DividedByNumber), Decimal_Precision, MidpointRounding.AwayFromZero);
            }
            catch (Exception ex)
            {
                ProcessError.Show("Global", "DivideNum", ex, new List<string>() { "Double Return", "DivNumber = " + DivNumber, "DividedByNumber = " + DividedByNumber, "Decimal_Precision = " + Decimal_Precision });
                return 0;
            }
        }

        public static DateTime ConvertToDateTime(string myDate)
        {
            try
            {
                // we need to have imported System.Globalization
                // using System.Globalization;

                // fetch the en-GB culture
                CultureInfo ukCulture = new CultureInfo("en-GB");
                // pass the DateTimeFormat information to DateTime.Parse
                if (myDate != null && myDate.Length > 1)
                {
                    DateTime myDateTime = DateTime.Parse(myDate, ukCulture.DateTimeFormat);
                    return myDateTime;
                }
                else
                {
                    return DateTime.Parse("01/01/1900");
                }
            }
            catch (Exception ex)
            {
                ProcessError.Show("Global", "ConvertToDateTime", ex, new List<string>() { "myDate = " + myDate });
                return DateTime.Parse("01/01/1900");
            }
        }

        public static DateTime ConvertToDateTime(object oDate)
        {
            try
            {
                // we need to have imported System.Globalization
                // using System.Globalization;
                string myDate = ConvertToString(oDate);
                // fetch the en-GB culture
                CultureInfo ukCulture = new CultureInfo("en-GB");
                // pass the DateTimeFormat information to DateTime.Parse
                if (myDate != null && myDate.Length > 1)
                {
                    DateTime myDateTime = DateTime.Parse(myDate, ukCulture.DateTimeFormat);
                    return myDateTime;
                }
                else
                {
                    return DateTime.Parse("01/01/1900");
                }
            }
            catch (Exception ex)
            {
                ProcessError.Show("Global", "ConvertToDateTime", ex, new List<string>() { "value = " + oDate });
                return DateTime.Parse("01/01/1900");
            }
        }

        public static double ConvertToDouble(object oValue)
        {
            try
            {
                string strValue = ConvertToString(oValue);
                if (string.IsNullOrEmpty(strValue))
                    return 0;

                return double.Parse(strValue);
            }
            catch
            {
                return 0;
            }
        }

        public static double ConvertToDouble(object oValue, int Decimal_Places)
        {
            try
            {
                double rValue = 0;

                if (Decimal_Places >= 0)
                    rValue = Math.Round(double.Parse(oValue.ToString()), Decimal_Places, MidpointRounding.AwayFromZero);
                else
                    rValue = double.Parse(oValue.ToString());

                return rValue;
            }
            catch
            {
                return 0;
            }
        }
    }
}
