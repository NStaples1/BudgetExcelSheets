using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.Classes
{
   internal static class Global
   {
      internal static string ConvertToString(object sValue)
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

      internal static string ConvertDoubleToString(double? sValue, string Formatting)
      {
         try
         {
            if (sValue != null)
            {
               if (sValue.HasValue)
                  return sValue.Value.ToString(Formatting);
               else
                  return string.Empty;
            }
            else
               return string.Empty;
         }
         catch
         {
            return string.Empty;
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

   }
}
