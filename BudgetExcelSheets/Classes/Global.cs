using DXTools;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BudgetExcelSheets.Classes
{
    public static class Global
    {
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
