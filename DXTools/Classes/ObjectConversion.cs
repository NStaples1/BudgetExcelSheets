using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.Classes
{
   internal static class ObjectConversion
   {
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
            ProcessError.Show("Global", "ConvertToDateTime", ex);
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
            ProcessError.Show("Global", "ConvertToDateTime", ex);
            return DateTime.Parse("01/01/1900");
         }
      }


      internal static double ConvertToDouble(object value)
      {
         try
         {
            if (value == DBNull.Value || value == null)
               return 0;
            else
               return double.Parse(value.ToString());
         }
         catch (Exception ex)
         {
            ProcessError.Show("Global", "CDBLCheck", ex);
            return 0;
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

      public static string ConvertDateToString(object sValue, string Format)
      {
         try
         {
            if (sValue != null)
            {
               DateTime RString = DateTime.Parse(sValue.ToString());
               return RString.ToString(Format);
            }
            else
               return string.Empty;
         }
         catch
         {
            return string.Empty;
         }
      }

      public static decimal ConvertToDecimal(object cValue, int DecimalPlaces)
      {
         try
         {
            return Math.Round(ConvertToDecimal(cValue), DecimalPlaces);
         }
         catch
         {
            return 0;
         }
      }

      public static int ConvertToInt(object cValue)
      {
         try
         {
            if (cValue == null || cValue == DBNull.Value)
               return 0;
            else
               return int.Parse(ConvertToString(cValue));
         }
         catch
         {
            return 0;
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

      public static bool ConvertToBool(object cValue)
      {
         try
         {
            if (cValue == null) return false;

            string strValue = cValue.ToString().ToLower();

            switch (strValue)
            {
               case "y":
               case "yes":
               case "t":
               case "true":
               case "1":
                  return true;
               case "n":
               case "no":
               case "f":
               case "false":
               case "0":
                  return false;
               default:
                  return false;
            }
         }
         catch
         {
            return false;
         }
      }
   }
}
