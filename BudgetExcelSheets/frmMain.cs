using BudgetExcelSheets.Classes;
using BudgetExcelSheets.Models;
using DevExpress.Charts.Native;
using DevExpress.Data.Helpers;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using DevExpress.Utils.About;
using DevExpress.Utils.Svg;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Filtering.Templates;
using DevExpress.XtraEditors.Popup;
using DevExpress.XtraExport.Implementation;
using DevExpress.XtraSpreadsheet.Import.Xls;
using DevExpress.XtraSpreadsheet.Model.CopyOperation;
using DevExpress.XtraSpreadsheet.Utils.Trees;
using DXTools;
using DXTools.Classes;
using SaltireAPI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VPS.Core;

namespace BudgetExcelSheets
{
   public partial class frmMain : DevExpress.XtraEditors.XtraForm
   {
      private const string ModuleName = "BudgetExcelSheets.frmMain";
      private SamsManagementSheet sms = new SamsManagementSheet();

      public frmMain()
      {
         InitializeComponent();
      }

      private void cmdGo_Click(object sender, EventArgs e)
      {
         CreateSpreadsheet();
      }

      internal void CreateSpreadsheet()
      {
         if (dteReportDate == null)
         {
            dteReportDate = new DevExpress.XtraEditors.DateEdit();
            dteReportDate.EditValue = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-01");
         }
         else
         {
            if (dteReportDate.EditValue == null)
               dteReportDate.EditValue = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-01");
         }

         using (DXTools.Spreadsheet sSheet = new DXTools.Spreadsheet())
         {
            List<string> NewnoBudgetList = new List<string>();
            List<string> BudgetNameList = new List<string>();
            List<string> OutdatedList = new List<string>();
            Color LightGreen = ColorTranslator.FromHtml("#66FFCC");
            int Top15 = 0;
            int NewBvBudget = 0;
            int TotalSalesvPriorYears = 0;
            int CustomervBdgvPy = 0;
            int ThisYearBudget = 0;
            int LastYearBudget = 0;
            int PriorYearBudget = 0;
            int ForecastWeekly = 0;
            int NewBusinessNoBudget = 0;
            int CurrentMonthTurnover = 0;
            int YTDSalesSheet = 0;
            int SalesvBudget = 0;
            int ThisYearMonthSalesPerCustomer = 0;
            int LastYearMonthSalesPerCustomer = 0;
            int PriorYearMonthSalesPerCustomer = 0;
            int ThisYearvsLastYear = 0;
            int CustomersNotBoughtThisMonth = 0;
            try
            {
               if (!Classes.Global.hasArgs)
                  sSheet.Show_Wait();


               string StartDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).ToString("yyyy-MM-dd");
               string EndDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddMonths(1).AddSeconds(-1).ToString("yyyy-MM-dd");
               string LastYearStartDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy-MM-dd");
               string LastYearStart = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy-01-01");
               string Year = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).ToString("yyyy");
               string LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy");
               string Month = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).ToString("MMMM").ToUpper();
               var MonthNo = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).Month;
               clsInvoices Invoices = new clsInvoices();
               int RowNumber = 0;
               int SheetNumber = 0;
               string sqlstring = string.Empty;
               /**************************************************************************************************************************
               * BUDGET
               *************************************************************************************************************************/

               sSheet.LoadFromFile("H:\\INTERNAL SALES\\Management Budget Template\\Budget Sheet.xlsx");

               sSheet.Set_Workbook_Units(Spreadsheet.DocumentUnits.Point);

               int MonthColumnIndex = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).Month + 1;
               int BudgetTotalRow = sSheet.GetWorksheetRange(Year + " BUDGET").RowCount;
               int priorYearBudgetTotalRow = sSheet.GetWorksheetRange(LastYear + " BUDGET").RowCount;
               int NewBusinessWonStart = 0;
               int NewBusinessWonEnd = 0;
               int newBusinessStart = 0;
               int newBusinessEnd = 0;
               SheetNumber = sSheet.Get_Worksheet_Index(Year + " BUDGET");
               ThisYearBudget = sSheet.Get_Worksheet_Index(Year + " BUDGET");
               LastYearBudget = sSheet.Get_Worksheet_Index(LastYear + " BUDGET");
               ForecastWeekly = sSheet.Get_Worksheet_Index("FORECAST WITH WEIGHT");
               int ForecastTotalRows = sSheet.GetWorksheetRange("FORECAST WITH WEIGHT").RowCount;

               List<BudgetModel> BudgetList = ConvertBudgetSheetToDataTable(SheetNumber, BudgetTotalRow, sSheet);

               RowNumber++;

               for (int i = 1; i < BudgetTotalRow; i++)
               {
                  sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (i + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (i + 1) + ")", 0, "#,##0");

                  if (Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 0, SheetNumber)) == "NEW BUSINESS WON IN " + LastYear + " IMPACTING " + Year ||
                     Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 0, SheetNumber)) == "NEW BUSINESS IN " + LastYear ||
                     Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 0, SheetNumber)) == "NEW BUSINESS WON IN " + LastYear)
                  {
                     NewBusinessWonStart = i + 1;
                  }
                  if (Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 0, SheetNumber)) == "NEW BUSINESS WON IN " + Year || Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 0, 0)) == "NEW BUSINESS IN " + Year)
                  {
                     NewBusinessWonEnd = i - 1;
                     newBusinessStart = i + 1;
                  }
                  if (Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 0, SheetNumber)) == "New / OTHER")
                     newBusinessEnd = i - 1;

                  RowNumber++;
               }

               RowNumber = 4;
               for (int i = 4; i < ForecastTotalRows - 1; i++)
               {
                  SetForecastYTD(sSheet, MonthColumnIndex, RowNumber, i);
                  RowNumber++;
               }

               for (int i = 1; i < (NewBusinessWonStart - 1); i++)
               {
                  string value = sSheet.Get_Cell_Text(i, 0, Year + " BUDGET");
                  if (value != null)
                     BudgetNameList.Add(value);
               }

               for (int i = NewBusinessWonStart; i < NewBusinessWonEnd; i++)
               {
                  string value = sSheet.Get_Cell_Text(i, 0, Year + " BUDGET");
                  if (value != null)
                     BudgetNameList.Add(value);
               }

               for (int i = newBusinessStart; i < newBusinessEnd; i++)
               {
                  string value = sSheet.Get_Cell_Text(i, 0, Year + " BUDGET");
                  if (value != null)
                     BudgetNameList.Add(value);
               }

               SheetNumber++;
               RowNumber = 0;

               /**************************************************************************************************************************
               * NEW BUSINESS NO BUDGET
               *************************************************************************************************************************/
               SheetNumber = (sSheet.Get_Sheet_Count() - 1);
               SheetNumber++;
               RowNumber = 0;
               sqlstring = "SELECT tbl_Customer.Name AS Name " +
               "FROM tbl_Customer LEFT OUTER JOIN " +
                   "(SELECT COUNT(tbl_Customer.Account_Ref) AS Account, tbl_Customer.Name, tbl_Customer.CustomerID " +
                   "FROM tbl_Customer INNER JOIN " +
                   "tbl_Invoice ON tbl_Customer.CustomerID = tbl_Invoice.CustomerID " +
                   "WHERE (tbl_Invoice.Invoice_Date BETWEEN CONVERT(DATETIME, '" + LastYear + "-01-01 00:00:00', 102) AND CONVERT(DATETIME, '" + LastYear + "-12-31 00:00:00', 102)) AND(tbl_Customer.Deleted = 0) " +
                   "GROUP BY tbl_Customer.Name, tbl_Customer.CustomerID) InvoicesPrevYear ON tbl_Customer.CustomerID = InvoicesPrevYear.CustomerID LEFT OUTER JOIN " +
                   "(SELECT COUNT(tbl_Customer.Account_Ref) AS Account, tbl_Customer.Name, tbl_Customer.CustomerID " +
                   "FROM tbl_Customer INNER JOIN " +
                   "tbl_Invoice ON tbl_Customer.CustomerID = tbl_Invoice.CustomerID " +
                   "WHERE(tbl_Invoice.Invoice_Date BETWEEN CONVERT(DATETIME, '" + Year + "-01-01 00:00:00', 102) AND CONVERT(DATETIME, '" + Year + "-12-31 00:00:00', 102)) AND(tbl_Customer.Deleted = 0) " +
                   "GROUP BY tbl_Customer.Name, tbl_Customer.CustomerID) InvoicesThisYear ON tbl_Customer.CustomerID = InvoicesThisYear.CustomerID " +
                   "WHERE InvoicesPrevYear.CustomerID IS NULL AND NOT(InvoicesThisYear.CustomerID IS NULL) ";


               //sqlstring = "SELECT DISTINCT tbl_Customer.Name " +
               //"FROM tbl_Customer INNER JOIN " +
               //"tbl_Invoice ON tbl_Customer.CustomerID = tbl_Invoice.CustomerID " +
               //"WHERE (tbl_Customer.Deleted = 0) AND  (tbl_Invoice.Invoice_Date BETWEEN CONVERT(DATETIME, '" + LastYear + "-01-01 00:00:00', 102) AND CONVERT(DATETIME, '" + Year + "-12-31 00:00:00', 102))";

               DataTable NewBusinessNoBudgetTable = Invoices.RetrieveDataTable(sqlstring, false);

               sSheet.Add_Worksheet("NEW BUSINESS NO BUDGET");
               NewBusinessNoBudget = SheetNumber;

               foreach (DataRow row in NewBusinessNoBudgetTable.Rows)
               {
                  if (row["Name"].ToString() == "STANNAH STAIRLIFT EURO ACCOUNT")
                     continue;
                  else
                  {
                     bool AddName = true;
                     for (int i = newBusinessStart; i < BudgetTotalRow; i++)
                     {
                        if (row["Name"].ToString() == sSheet.Get_Cell_Text(i, 0, Year + " BUDGET"))
                           AddName = false;
                     }
                     if (AddName)
                     {
                        sSheet.Set_Cell(RowNumber, 0, row["Name"].ToString(), SheetNumber);

                        RowNumber++;
                     }
                  }
               }

               int newCustomerNoBudgetTotalRows = sSheet.GetWorksheetRange("NEW BUSINESS NO BUDGET").RowCount;

               NewBusinessNoBudgetTable.Dispose();
               NewBusinessNoBudgetTable = null;

               sSheet.Auto_fit(0, 1, SheetNumber);

               for (int i = 0; i < newCustomerNoBudgetTotalRows; i++)
               {
                  string value = sSheet.Get_Cell_Text(i, 0, "NEW BUSINESS NO BUDGET");
                  if (value != null)
                  {
                     if (!BudgetNameList.Contains(value))
                        NewnoBudgetList.Add(value);
                  }
               }

               /**************************************************************************************************************************
               * CURRENT MONTH TURNOVER SUMMARY 
               * Current month is equal to previous month
               *************************************************************************************************************************/

               RowNumber = 0;
               SheetNumber++;

               sqlstring = "SELECT tbl_Invoice.CustomerID, SUM(tbl_InvoiceItem.Cost_Price * tbl_InvoiceItem.Qty_Order) AS Line_Cost_Price, SUM(tbl_InvoiceItem.Net_Amount) AS Line_Sale_Price, " +
                           "tbl_Customer.Name, tbl_Customer.Account_Ref, " +
                           "SUM(tbl_Product.Unit_Weight * tbl_InvoiceItem.Qty_Order) AS Line_Unit_Weight " +
                           "FROM tbl_Invoice AS tbl_Invoice INNER JOIN " +
                           "tbl_InvoiceItem ON tbl_Invoice.InvoiceID = tbl_InvoiceItem.InvoiceID LEFT OUTER JOIN " +
                           "tbl_Product ON tbl_InvoiceItem.ProductID = tbl_Product.ProductID LEFT OUTER JOIN " +
                           "tbl_Customer AS tbl_Customer ON tbl_Invoice.CustomerID = tbl_Customer.CustomerID " +
                           "WHERE(tbl_Invoice.Invoice_Date BETWEEN '" + StartDate + "' AND '" + EndDate + "') " +
                           "GROUP BY tbl_Invoice.CustomerID, tbl_Customer.Name, tbl_Customer.Account_Ref " +
                           "ORDER BY tbl_Customer.Name ";

               DataTable MonthTurnoverTable = Invoices.RetrieveDataTable(sqlstring, false);

               sSheet.Add_Worksheet("CURRENT MONTH TURNOVER SUMMARY");
               CurrentMonthTurnover = SheetNumber;

               sSheet.Set_Cell(RowNumber, 0, "Customer Name", SheetNumber);
               sSheet.Set_Cell(RowNumber, 1, "Cost Price", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "Sales Price", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "Profit", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "AVG M/U%", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "AVG P/S%", SheetNumber);

               RowNumber++;

               foreach (DataRow row in MonthTurnoverTable.Rows)
               {
                  double Profit = 0;
                  if (row["Name"].ToString() == "STANNAH STAIRLIFT EURO ACCOUNT")
                  {
                     DataRow[] EuroRow = MonthTurnoverTable.Select("Name = 'STANNAH STAIRLIFTS LTD'");
                     double EuroCost = 0;
                     double EuroSale = 0;

                     if (EuroRow.Length > 0)
                     {
                        EuroSale = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Sale_Price"], 2);
                        EuroCost = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Cost_Price"], 2);
                     }

                     Profit = (Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2)) + EuroSale - EuroCost;

                     sSheet.Set_Cell(RowNumber, 0, "STANNAH STAIRLIFTS LTD", SheetNumber);
                     sSheet.Set_Cell(RowNumber, 1, Classes.Global.ConvertToDouble(row["Line_Cost_Price"]) + EuroCost, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 2, Classes.Global.ConvertToDouble(row["Line_Sale_Price"]) + EuroSale, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalCost = Classes.Global.ConvertToDouble(row["Line_Cost_Price"]) + EuroCost;
                     double ProfitTotalCost = Classes.Global.DivideNum(Profit, TotalCost, 4);
                     double ProfitMarginCostPercentage = ProfitTotalCost * 100;

                     sSheet.Set_Cell(RowNumber, 4, ProfitMarginCostPercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalSale = Classes.Global.ConvertToDouble(row["Line_Sale_Price"]) + EuroSale;
                     double ProfitTotalSale = Classes.Global.DivideNum(Profit, TotalSale, 4);
                     double ProfitMarginSalePercentage = ProfitTotalSale * 100;

                     sSheet.Set_Cell(RowNumber, 5, ProfitMarginSalePercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     RowNumber++;
                  }
                  else if (row["Name"].ToString() == "STANNAH STAIRLIFTS LTD")
                  {

                  }
                  else if (row["Name"].ToString() == "DIGICO (UK) LTD")
                  {
                     DataRow[] EuroRow = MonthTurnoverTable.Select("Name = 'A6 AUDIO LIMITED'");
                     double EuroCost = 0;
                     double EuroSale = 0;

                     if (EuroRow.Length > 0)
                     {
                        EuroSale = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Sale_Price"], 2);
                        EuroCost = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Cost_Price"], 2);
                     }

                     Profit = (Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2)) + EuroSale - EuroCost;

                     sSheet.Set_Cell(RowNumber, 0, "DIGICO (UK) LTD", SheetNumber);
                     sSheet.Set_Cell(RowNumber, 1, Classes.Global.ConvertToDouble(row["Line_Cost_Price"]) + EuroCost, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 2, Classes.Global.ConvertToDouble(row["Line_Sale_Price"]) + EuroSale, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalCost = Classes.Global.ConvertToDouble(row["Line_Cost_Price"]) + EuroCost;
                     double ProfitTotalCost = Classes.Global.DivideNum(Profit, TotalCost, 4);
                     double ProfitMarginCostPercentage = ProfitTotalCost * 100;

                     sSheet.Set_Cell(RowNumber, 4, ProfitMarginCostPercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalSale = Classes.Global.ConvertToDouble(row["Line_Sale_Price"]) + EuroSale;
                     double ProfitTotalSale = Classes.Global.DivideNum(Profit, TotalSale, 4);
                     double ProfitMarginSalePercentage = ProfitTotalSale * 100;

                     sSheet.Set_Cell(RowNumber, 5, ProfitMarginSalePercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     RowNumber++;
                  }
                  else if (row["Name"].ToString() == "A6 AUDIO LIMITED")
                  {

                  }
                  else
                  {
                     Profit = Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2);
                     sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]).Trim(), SheetNumber);
                     sSheet.Set_Cell(RowNumber, 1, row["Line_Cost_Price"], SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 2, row["Line_Sale_Price"], SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 4, Classes.Global.DivideNum(Profit, Classes.Global.ConvertToDouble(row["Line_Cost_Price"]), 4) * 100, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 5, Classes.Global.DivideNum(Profit, Classes.Global.ConvertToDouble(row["Line_Sale_Price"]), 4) * 100, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     RowNumber++;
                  }
               }
               sSheet.FormatCell("F2:F" + (RowNumber + 1), "#,##0", SheetNumber);

               sSheet.Set_Formula(RowNumber + 1, 2, "=SUBTOTAL(9,C2:C" + (RowNumber) + ")", SheetNumber, "#,##0.00");
               sSheet.Set_Formula(RowNumber + 1, 3, "=SUBTOTAL(9,D2:D" + (RowNumber) + ")", SheetNumber, "#,##0.00");

               int MonthlyTotalRowNumber = RowNumber + 2;

               sSheet.Auto_fit(0, 5, SheetNumber);

               /**************************************************************************************************************************
               * YTD SALES
               *************************************************************************************************************************/

               RowNumber = 0;
               SheetNumber++;

               //sqlstring = "SELECT tbl_Invoice.CustomerID, SUM(tbl_InvoiceItem.Cost_Price * tbl_InvoiceItem.Qty_Order) AS Line_Cost_Price, SUM(tbl_InvoiceItem.Net_Amount) AS Line_Sale_Price, " +
               //            "tbl_Customer.Name, tbl_Customer.Account_Ref, " +
               //            "SUM(tbl_Product.Unit_Weight * tbl_InvoiceItem.Qty_Order) AS Line_Unit_Weight " +
               //            "FROM tbl_Invoice AS tbl_Invoice INNER JOIN " +
               //            "tbl_InvoiceItem ON tbl_Invoice.InvoiceID = tbl_InvoiceItem.InvoiceID LEFT OUTER JOIN " +
               //            "tbl_Product ON tbl_InvoiceItem.ProductID = tbl_Product.ProductID LEFT OUTER JOIN " +
               //            "tbl_Customer AS tbl_Customer ON tbl_Invoice.CustomerID = tbl_Customer.CustomerID " +
               //            "WHERE(tbl_Invoice.Invoice_Date BETWEEN '" + Year + "-01-01' AND '" + EndDate + "') " +
               //            "GROUP BY tbl_Invoice.CustomerID, tbl_Customer.Name, tbl_Customer.Account_Ref " +
               //            "ORDER BY tbl_Customer.Account_Ref ";

               sqlstring = "SELECT tbl_Customer.CustomerID, tbl_Customer.Account_Ref,LTRIM(RTRIM(tbl_Customer.Name)) AS Name, Inv.Line_Cost_Price, Inv.Line_Sale_Price, Inv.Line_Unit_Weight " +
                   "FROM tbl_Customer LEFT OUTER JOIN(SELECT SUM(tbl_InvoiceItem.Cost_Price* tbl_InvoiceItem.Qty_Order) AS Line_Cost_Price, SUM(tbl_InvoiceItem.Net_Amount) AS Line_Sale_Price, " +
                   "SUM(tbl_Product.Unit_Weight * tbl_InvoiceItem.Qty_Order) AS Line_Unit_Weight, " +
                   "tbl_Invoice.CustomerID " +
                   "FROM tbl_Invoice AS tbl_Invoice LEFT OUTER JOIN " +
                   "tbl_Product RIGHT OUTER JOIN " +
                   "tbl_InvoiceItem ON tbl_Product.ProductID = tbl_InvoiceItem.ProductID ON tbl_Invoice.InvoiceID = tbl_InvoiceItem.InvoiceID " +
                   "WHERE(tbl_Invoice.Invoice_Date IS NULL OR " +
                   "tbl_Invoice.Invoice_Date BETWEEN '" + Year + "-01-01' AND '" + EndDate + "') " +
                   "GROUP BY tbl_Invoice.CustomerID) Inv ON tbl_Customer.CustomerID = Inv.CustomerID " +
                   "WHERE tbl_Customer.Deleted = 0 OR tbl_Customer.CustomerID = '6c009bb3-84d4-48cc-97cc-379cd017d3c7' " +
                   "ORDER BY tbl_Customer.Name ";

               DataTable YTDSales = Invoices.RetrieveDataTable(sqlstring, false);

               foreach (BudgetModel BudgetRecord in BudgetList.Where(w => w.Section == "NEW BUSINESS IN " + Year).ToList())
               {
                  DataRow[] checkExists = YTDSales.Select("Name = '" + BudgetRecord.ExistingCustomers + "'");
                  if (checkExists.Length == 0)
                  {
                     DataRow newCustomerRow = YTDSales.NewRow();
                     newCustomerRow["Name"] = BudgetRecord.ExistingCustomers;
                     YTDSales.Rows.Add(newCustomerRow);
                  }
               }

               YTDSales = resort(YTDSales, "Name", "ASC");

               sSheet.Add_Worksheet("YTD SALES");
               YTDSalesSheet = SheetNumber;

               sSheet.Set_Cell(RowNumber, 0, "Customer Name", SheetNumber);
               sSheet.Set_Cell(RowNumber, 1, "Cost Price", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "Sales Price", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "Profit", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "AVG M/U%", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "AVG P/S%", SheetNumber);

               RowNumber++;

               foreach (DataRow row in YTDSales.Rows)
               {
                  double Profit = 0;
                  if (row["Name"].ToString() == "STANNAH STAIRLIFT EURO ACCOUNT")
                  {
                     DataRow[] EuroRow = YTDSales.Select("Name = 'STANNAH STAIRLIFTS LTD'");
                     double EuroCost = 0;
                     double EuroSale = 0;

                     if (EuroRow.Length > 0)
                     {
                        EuroSale = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Sale_Price"], 2);
                        EuroCost = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Cost_Price"], 2);
                     }

                     double Cost = Classes.Global.ConvertToDouble(row["Line_Cost_Price"]);
                     double Sale = Classes.Global.ConvertToDouble(row["Line_Sale_Price"]);

                     Profit = (Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2)) + EuroSale - EuroCost;

                     sSheet.Set_Cell(RowNumber, 0, "STANNAH STAIRLIFTS LTD", SheetNumber);
                     sSheet.Set_Cell(RowNumber, 1, Cost + EuroCost, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 2, Sale + EuroSale, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalCost = Cost + EuroCost;
                     double ProfitTotalCost = Classes.Global.DivideNum(Profit, TotalCost, 4);
                     double ProfitMarginCostPercentage = ProfitTotalCost * 100;
                     sSheet.Set_Cell(RowNumber, 4, ProfitMarginCostPercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalSale = Sale + EuroSale;
                     double ProfitTotalSale = Classes.Global.DivideNum(Profit, TotalSale, 4);
                     double ProfitMarginSalePercentage = ProfitTotalSale * 100;
                     sSheet.Set_Cell(RowNumber, 5, ProfitMarginSalePercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     RowNumber++;
                  }
                  else if (row["Name"].ToString() == "STANNAH STAIRLIFTS LTD")
                  { }
                  else if (row["Name"].ToString() == "NEW WAVE DOORS DIRECT LTD")
                  {
                     DataRow[] EuroRow = YTDSales.Select("Name = 'DELTACO 1 LTD'");
                     double EuroCost = 0;
                     double EuroSale = 0;

                     if (EuroRow.Length > 0)
                     {
                        EuroSale = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Sale_Price"], 2);
                        EuroCost = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Cost_Price"], 2);
                     }

                     double Cost = Classes.Global.ConvertToDouble(row["Line_Cost_Price"]);
                     double Sale = Classes.Global.ConvertToDouble(row["Line_Sale_Price"]);

                     Profit = (Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2)) + EuroSale - EuroCost;

                     sSheet.Set_Cell(RowNumber, 0, "NEW WAVE DOORS DIRECT LTD", SheetNumber);
                     sSheet.Set_Cell(RowNumber, 1, Cost + EuroCost, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 2, Sale + EuroSale, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalCost = Cost + EuroCost;
                     double ProfitTotalCost = Classes.Global.DivideNum(Profit, TotalCost, 4);
                     double ProfitMarginCostPercentage = ProfitTotalCost * 100;
                     sSheet.Set_Cell(RowNumber, 4, ProfitMarginCostPercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalSale = Sale + EuroSale;
                     double ProfitTotalSale = Classes.Global.DivideNum(Profit, TotalSale, 4);
                     double ProfitMarginSalePercentage = ProfitTotalSale * 100;
                     sSheet.Set_Cell(RowNumber, 5, ProfitMarginSalePercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     RowNumber++;
                  }
                  else if (row["Name"].ToString() == "DELTACO 1 LTD")
                  { }
                  else if (row["Name"].ToString() == "DIGICO (UK) LTD")
                  {
                     DataRow[] EuroRow = YTDSales.Select("Name = 'A6 AUDIO LIMITED'");
                     double EuroCost = 0;
                     double EuroSale = 0;

                     if (EuroRow.Length > 0)
                     {
                        EuroSale = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Sale_Price"], 2);
                        EuroCost = Classes.Global.ConvertToDouble(EuroRow[0]["Line_Cost_Price"], 2);
                     }

                     double Cost = Classes.Global.ConvertToDouble(row["Line_Cost_Price"]);
                     double Sale = Classes.Global.ConvertToDouble(row["Line_Sale_Price"]);

                     Profit = (Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2)) + EuroSale - EuroCost;

                     sSheet.Set_Cell(RowNumber, 0, "DIGICO (UK) LTD", SheetNumber);
                     sSheet.Set_Cell(RowNumber, 1, Cost + EuroCost, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 2, Sale + EuroSale, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalCost = Cost + EuroCost;
                     double ProfitTotalCost = Classes.Global.DivideNum(Profit, TotalCost, 4);
                     double ProfitMarginCostPercentage = ProfitTotalCost * 100;
                     sSheet.Set_Cell(RowNumber, 4, ProfitMarginCostPercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     double TotalSale = Sale + EuroSale;
                     double ProfitTotalSale = Classes.Global.DivideNum(Profit, TotalSale, 4);
                     double ProfitMarginSalePercentage = ProfitTotalSale * 100;
                     sSheet.Set_Cell(RowNumber, 5, ProfitMarginSalePercentage, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     RowNumber++;
                  }
                  else if (row["Name"].ToString() == "A6 AUDIO LIMITED")
                  { }
                  else
                  {
                     Profit = Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2);
                     sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]).Trim(), SheetNumber);
                     sSheet.Set_Cell(RowNumber, 1, row["Line_Cost_Price"], SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 2, row["Line_Sale_Price"], SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 4, Classes.Global.DivideNum(Profit, Classes.Global.ConvertToDouble(row["Line_Cost_Price"]), 4) * 100, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, 5, Classes.Global.DivideNum(Profit, Classes.Global.ConvertToDouble(row["Line_Sale_Price"]), 4) * 100, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                     RowNumber++;
                  }
               }
               sSheet.FormatCell("F2:F" + (RowNumber + 1), "#,##0", SheetNumber);

               sSheet.Set_Formula(RowNumber + 1, 2, "=SUBTOTAL(9,C2:C" + (RowNumber) + ")", SheetNumber, "#,##0.00");
               sSheet.Set_Formula(RowNumber + 1, 3, "=SUBTOTAL(9,D2:D" + (RowNumber) + ")", SheetNumber, "#,##0.00");

               int YTDTotalRowNumber = RowNumber + 2;

               sSheet.Auto_fit(0, 5, SheetNumber);

               /**************************************************************************************************************************
                * DATA PRESENTATION
                * Already have the data in CURRENT MONTH
                * Sort datatable
                *************************************************************************************************************************/

               Color Colour = System.Drawing.ColorTranslator.FromHtml("#009999");

               SheetNumber++;
               RowNumber = 1;
               sSheet.Insert_Worksheet("1.TOP 15", SheetNumber);
               Top15 = SheetNumber;

               sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToDateTime(dteReportDate.EditValue).ToString("MMMM"), SheetNumber);
               sSheet.Set_Bold(RowNumber, 0, true, SheetNumber);
               sSheet.Merge_Cells("A" + (RowNumber + 1) + ":G" + (RowNumber + 1), SheetNumber);
               sSheet.Set_Font_Size("A" + (RowNumber + 1), 20, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":G" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               sSheet.Set_Cell(RowNumber, 8, "Year To Date", SheetNumber);
               sSheet.Set_Bold(RowNumber, 8, true, SheetNumber);
               sSheet.Merge_Cells("I" + (RowNumber + 1) + ":O" + (RowNumber + 1), SheetNumber);
               sSheet.Set_Font_Size("I" + (RowNumber + 1), 20, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("I" + (RowNumber + 1) + ":O" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               RowNumber++;
               RowNumber++;

               sSheet.Set_Row_Height(RowNumber, 45.00, SheetNumber);

               sSheet.Set_Font_Size("A" + (RowNumber + 1), 14, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":G" + (RowNumber + 1), Colour, Color.White, SheetNumber);
               sSheet.Set_Cell(RowNumber, 0, "CUSTOMER", SheetNumber, SpreadsheetHorizontalAlignment.Left, true);
               sSheet.Set_Cell(RowNumber, 1, "ACTUAL SALES £", SheetNumber, SpreadsheetHorizontalAlignment.Left, true);
               sSheet.Set_Cell(RowNumber, 2, "BUDGET SALES £", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 3, "SALES V BUDGET £", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 4, "SALES V BUDGET %", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 5, "MARGIN £", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 6, "MARGIN %", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);


               sSheet.Set_Cell(RowNumber, 8, "CUSTOMER", SheetNumber, SpreadsheetHorizontalAlignment.Left, true);
               sSheet.Set_Cell(RowNumber, 9, "ACTUAL SALES £", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 10, "BUDGET SALES £", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 11, "SALES V BUDGET £", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 12, "SALES V BUDGET %", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 13, "MARGIN £", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Cell(RowNumber, 14, "MARGIN %", SheetNumber, SpreadsheetHorizontalAlignment.Center, true);
               sSheet.Set_Font_Size("I" + (RowNumber + 1), 14, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("I" + (RowNumber + 1) + ":O" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               RowNumber++;
               int YTDSalesTop15 = RowNumber;
               sSheet.Set_Cell("H" + (RowNumber + 1), "Top 15 Customers", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold(RowNumber, 0, true, SheetNumber);
               sSheet.Set_Rotation("H" + (RowNumber + 1), SheetNumber, 90, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("H" + (RowNumber + 1), 14, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("H" + (RowNumber + 1) + ":H" + (RowNumber + 1), LightGreen, Color.Black, SheetNumber);

               DataRow[] Digico = MonthTurnoverTable.Select("Name = 'DIGICO (UK) LTD'");
               DataRow[] A6Audio = MonthTurnoverTable.Select("Name = 'A6 AUDIO LIMITED'");
               if (Digico.Length == 0)
               {
                  DataRow dataRow = MonthTurnoverTable.NewRow();
                  Digico = new DataRow[] { dataRow };
               }
               if (A6Audio.Length == 0)
               {
                  DataRow dataRow = MonthTurnoverTable.NewRow();
                  A6Audio = new DataRow[] { dataRow };
               }

               double CostPrice = Classes.Global.ConvertToDouble(Digico[0]["Line_Cost_Price"]) + Classes.Global.ConvertToDouble(A6Audio[0]["Line_Cost_Price"]);
               double SalePrice = Classes.Global.ConvertToDouble(Digico[0]["Line_Sale_Price"]) + Classes.Global.ConvertToDouble(A6Audio[0]["Line_Sale_Price"]);
               double UnitWeight = Classes.Global.ConvertToDouble(Digico[0]["Line_Unit_Weight"]) + Classes.Global.ConvertToDouble(A6Audio[0]["Line_Unit_Weight"]);

               Digico[0]["Line_Cost_Price"] = CostPrice;
               Digico[0]["Line_Sale_Price"] = SalePrice;
               Digico[0]["Line_Unit_Weight"] = UnitWeight;

               DataTable newMonthTable = resort(MonthTurnoverTable, "Line_Sale_Price", "DESC");

               DataRow[] NewWaveDoors = YTDSales.Select("Name = 'NEW WAVE DOORS DIRECT LTD'");
               DataRow[] DelTaco = YTDSales.Select("Name = 'DELTACO 1 LTD'");
               if(NewWaveDoors.Length == 0)
               {
                  DataRow dataRow = YTDSales.NewRow();
                  NewWaveDoors = new DataRow[] { dataRow };
               }
               if(DelTaco.Length == 0)
               {
                  DataRow dataRow = YTDSales.NewRow();
                  DelTaco = new DataRow[] { dataRow };
               }
               Digico = YTDSales.Select("Name = 'DIGICO (UK) LTD'");
               A6Audio = YTDSales.Select("Name = 'A6 AUDIO LIMITED'");
               if(Digico.Length == 0)
               {
                  DataRow dataRow = YTDSales.NewRow();
                  Digico = new DataRow[] { dataRow };
               }
               if(A6Audio.Length == 0)
               {
                  DataRow dataRow = YTDSales.NewRow();
                  A6Audio = new DataRow[] { dataRow };
               }

               CostPrice = Classes.Global.ConvertToDouble(NewWaveDoors[0]["Line_Cost_Price"]) + Classes.Global.ConvertToDouble(DelTaco[0]["Line_Cost_Price"]);
               SalePrice = Classes.Global.ConvertToDouble(NewWaveDoors[0]["Line_Sale_Price"]) + Classes.Global.ConvertToDouble(DelTaco[0]["Line_Sale_Price"]);
               UnitWeight = Classes.Global.ConvertToDouble(NewWaveDoors[0]["Line_Unit_Weight"]) + Classes.Global.ConvertToDouble(DelTaco[0]["Line_Unit_Weight"]);

               NewWaveDoors[0]["Line_Cost_Price"] = CostPrice;
               NewWaveDoors[0]["Line_Sale_Price"] = SalePrice;
               NewWaveDoors[0]["Line_Unit_Weight"] = UnitWeight;

               CostPrice = Classes.Global.ConvertToDouble(Digico[0]["Line_Cost_Price"]) + Classes.Global.ConvertToDouble(A6Audio[0]["Line_Cost_Price"]);
               SalePrice = Classes.Global.ConvertToDouble(Digico[0]["Line_Sale_Price"]) + Classes.Global.ConvertToDouble(A6Audio[0]["Line_Sale_Price"]);
               UnitWeight = Classes.Global.ConvertToDouble(Digico[0]["Line_Unit_Weight"]) + Classes.Global.ConvertToDouble(A6Audio[0]["Line_Unit_Weight"]);

               Digico[0]["Line_Cost_Price"] = CostPrice;
               Digico[0]["Line_Sale_Price"] = SalePrice;
               Digico[0]["Line_Unit_Weight"] = UnitWeight;

               DataTable newYTDTable = resort(YTDSales, "Line_Sale_Price", "DESC");

               for (int i = 0; i < 15; i++)
               {
                  if (newMonthTable.Rows[i]["Name"].ToString() == "STANNAH STAIRLIFTS LTD")
                  { }
                  else
                  {
                     if (newMonthTable.Rows[i]["Name"].ToString() == "STANNAH STAIRLIFT EURO ACCOUNT")
                        sSheet.Set_Cell(RowNumber, 0, "STANNAH STAIRLIFTS LTD", SheetNumber);
                     else if (newMonthTable.Rows[i]["Name"].ToString() == "A6 AUDIO LIMITED")
                        sSheet.Set_Cell(RowNumber, 0, "DIGICO (UK) LTD", SheetNumber);
                     else
                        sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(newMonthTable.Rows[i]["Name"]).Trim(), SheetNumber);


                     sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:C,3,FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O," + MonthColumnIndex + ",FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 3, "=" + sSheet.GetExcelColumnName(2) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(3) + (RowNumber + 1), SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 4, "=IFERROR(" + sSheet.GetExcelColumnName(4) + (RowNumber + 1) + "/" + sSheet.GetExcelColumnName(3) + (RowNumber + 1) + ",\"NO BUDGET\")", SheetNumber, "#,##0 %");
                     sSheet.Set_Formula(RowNumber, 5, "=VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!$A:D,4,FALSE)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 6, "=VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!$A:F,6,FALSE)", SheetNumber, "#,##0");
                     RowNumber++;
                  }
               }

               RowNumber = YTDSalesTop15;

               for (int i = 0; i < 15; i++)
               {
                  if (newYTDTable.Rows[i]["Name"].ToString() == "STANNAH STAIRLIFTS LTD")
                  { }
                  else
                  {
                     if (newYTDTable.Rows[i]["Name"].ToString() == "STANNAH STAIRLIFT EURO ACCOUNT")
                        sSheet.Set_Cell(RowNumber, 8, "STANNAH STAIRLIFTS LTD", SheetNumber);
                     else if (newYTDTable.Rows[i]["Name"].ToString() == "A6 AUDIO LIMITED")
                        sSheet.Set_Cell(RowNumber, 8, "DIGICO (UK) LTD", SheetNumber);
                     else
                        sSheet.Set_Cell(RowNumber, 8, Classes.Global.ConvertToString(newYTDTable.Rows[i]["Name"]).Trim(), SheetNumber);

                     sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(9) + (RowNumber + 1) + ",'YTD SALES'!A:C,3,FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 10, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(9) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O,14,FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 11, "=" + sSheet.GetExcelColumnName(10) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(11) + (RowNumber + 1), SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 12, "=IFERROR(" + sSheet.GetExcelColumnName(12) + (RowNumber + 1) + "/" + sSheet.GetExcelColumnName(11) + (RowNumber + 1) + ",\"NO BUDGET\")", SheetNumber, "#,##0 %");
                     sSheet.Set_Formula(RowNumber, 13, "=VLOOKUP(" + sSheet.GetExcelColumnName(9) + (RowNumber + 1) + ",'YTD SALES'!$A:D,4,FALSE)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 14, "=VLOOKUP(" + sSheet.GetExcelColumnName(9) + (RowNumber + 1) + ",'YTD SALES'!$A:F,6,FALSE)", SheetNumber, "#,##0");
                     RowNumber++;
                  }
               }

               sSheet.Merge_Cells("H" + (RowNumber - 14) + ":H" + (RowNumber), SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, "TOTAL OF TOP 15", SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Formula(RowNumber, 1, "=SUM(B" + (RowNumber - 15) + ":B" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(C" + (RowNumber - 15) + ":C" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=B" + (RowNumber + 1) + "-C" + (RowNumber + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=D" + (RowNumber + 1) + "/C" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=SUM(F" + (RowNumber - 15) + ":F" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=F" + (RowNumber + 1) + "/B" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Font_Size("A" + (RowNumber + 1), 12, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":G" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               sSheet.Set_Cell(RowNumber, 8, "TOTAL OF TOP 15", SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Formula(RowNumber, 9, "=SUM(J" + (RowNumber - 15) + ":J" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=SUM(K" + (RowNumber - 15) + ":K" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=J" + (RowNumber + 1) + "-K" + (RowNumber + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=L" + (RowNumber + 1) + "/K" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(N" + (RowNumber - 15) + ":N" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 14, "=N" + (RowNumber + 1) + "/J" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Font_Size("I" + (RowNumber + 1), 12, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("I" + (RowNumber + 1) + ":O" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               RowNumber++;
               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "TOTAL OF OTHERS", SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Formula(RowNumber, 1, "='CURRENT MONTH TURNOVER SUMMARY'!C" + (MonthlyTotalRowNumber) + "-'1.TOP 15'!B" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + Year + " BUDGET'!" + sSheet.GetExcelColumnName(MonthColumnIndex) + BudgetTotalRow + " - C" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=B" + (RowNumber + 1) + "-C" + (RowNumber + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=D" + (RowNumber + 1) + "/C" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "='CURRENT MONTH TURNOVER SUMMARY'!D" + (MonthlyTotalRowNumber) + "-F" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=F" + (RowNumber + 1) + "/B" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Font_Size("A" + (RowNumber + 1), 12, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":G" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               sSheet.Set_Cell(RowNumber, 8, "TOTAL OF OTHERS", SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Formula(RowNumber, 9, "='YTD SALES'!C" + (YTDTotalRowNumber) + "-'1.TOP 15'!J" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + Year + " BUDGET'!N" + BudgetTotalRow + " - K" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=J" + (RowNumber + 1) + "-K" + (RowNumber + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=L" + (RowNumber + 1) + "/K" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 13, "='YTD SALES'!D" + (YTDTotalRowNumber) + "-N" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 14, "=N" + (RowNumber + 1) + "/J" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Font_Size("I" + (RowNumber + 1), 12, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("I" + (RowNumber + 1) + ":O" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               RowNumber++;
               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "TOTAL OF ALL CUSTOMERS", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Formula(RowNumber, 1, "=B" + (RowNumber - 3) + "+B" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (RowNumber - 3) + "+C" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=B" + (RowNumber + 1) + "-C" + (RowNumber + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=D" + (RowNumber + 1) + "/C" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=F" + (RowNumber - 3) + "+F" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=F" + (RowNumber + 1) + "/B" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Font_Size("A" + (RowNumber + 1), 12, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":G" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               sSheet.Set_Cell(RowNumber, 8, "TOTAL OF ALL CUSTOMERS", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Formula(RowNumber, 9, "=J" + (RowNumber - 3) + "+J" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=K" + (RowNumber - 3) + "+K" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=J" + (RowNumber + 1) + "-K" + (RowNumber + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=L" + (RowNumber + 1) + "/K" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 13, "=N" + (RowNumber - 3) + "+N" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 14, "=N" + (RowNumber + 1) + "/J" + (RowNumber + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Font_Size("I" + (RowNumber + 1), 12, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_FontColour("I" + (RowNumber + 1) + ":O" + (RowNumber + 1), Colour, Color.White, SheetNumber);

               sSheet.Set_Cell_Alignment("B" + (RowNumber - 19) + ":G" + (RowNumber + 1), "1.TOP 15", SpreadsheetHorizontalAlignment.Center, SpreadsheetVerticalAlignment.Center, true);
               sSheet.Set_Cell_Alignment("J" + (RowNumber - 19) + ":O" + (RowNumber + 1), "1.TOP 15", SpreadsheetHorizontalAlignment.Center, SpreadsheetVerticalAlignment.Center, true);

               newMonthTable.Dispose();
               newMonthTable = null;
               newYTDTable.Dispose();
               newYTDTable = null;

               sSheet.Set_OutsideBorders("H" + (RowNumber - 18) + ":H" + (RowNumber - 4), Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);

               RowNumber += 2;

               sSheet.Set_Formula(RowNumber, 1, "=VLOOKUP(\"STANNAH STAIRLIFTS LTD\",A:B,2,FALSE)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=B20 - B" + (RowNumber + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=B22", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=SUM(B" + (RowNumber + 1) + ":D" + (RowNumber + 1) + ")", SheetNumber, "£#,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 1, "STANNAH", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "TOP 14", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "OTHERS", SheetNumber);
               sSheet.Set_Bold_Range("B" + (RowNumber + 1) + ":D" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;

               sSheet.Set_Formula(RowNumber, 1, "=B" + (RowNumber - 1) + "/E" + (RowNumber - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (RowNumber - 1) + " / E" + (RowNumber - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 3, "=D" + (RowNumber - 1) + " / E" + (RowNumber - 1), SheetNumber, "#,##0 %");

               sms.Set_Pie_Chart(sSheet, "B27:D28", DevExpress.Spreadsheet.Charts.ChartType.Pie, SheetNumber, "G" + (RowNumber - 1), "I" + (RowNumber + 10), "SALES RATIO: STANNAH V OTHERS");

               sSheet.Set_AllBorders("A2:G2", Color.Black, DevExpress.Spreadsheet.BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("I2:O2", Color.Black, DevExpress.Spreadsheet.BorderLineStyle.Thin, SheetNumber);

               sSheet.Set_AllBorders("A5:G19", Color.Black, DevExpress.Spreadsheet.BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("I5:O19", Color.Black, DevExpress.Spreadsheet.BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_OutsideBorders("A4:G24", Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("I4:O24", Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);

               sSheet.Set_OutsideBorders("A20:G20", Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("I20:O20", Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);

               sSheet.Set_OutsideBorders("A22:G22", Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("I22:O22", Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);

               sSheet.Set_OutsideBorders("A24:G24", Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("I24:O24", Color.Black, SheetNumber, DevExpress.Spreadsheet.BorderLineStyle.Medium);

               sSheet.Auto_fit(0, 15, SheetNumber);
               sSheet.Set_Column_Width("A", 33.57, "1.TOP 15");
               sSheet.Set_Column_Width("I", 33.57, "1.TOP 15");

               /**************************************************************************************************************************
               * SALES V BUDGET
               *************************************************************************************************************************/

               SheetNumber++;
               RowNumber = 0;
               sSheet.Insert_Worksheet("SALES V BUDGET", SheetNumber);
               SalesvBudget = SheetNumber;

               sSheet.Set_Cell(RowNumber, 0, "CUSTOMER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Bold("A" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_Cell(RowNumber, 1, "CURRENT MONTH SALES", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold("B" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_Cell(RowNumber, 2, "CURRENT MONTH BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold("C" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_Cell(RowNumber, 3, "YTD SALES", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold("D" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_Cell(RowNumber, 4, "YTD BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold("E" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_Cell(RowNumber, 6, "CURRENT MONTH MARGIN £", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold("G" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_Cell(RowNumber, 7, "CURRENT MONTH MARGIN %", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold("H" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_Cell(RowNumber, 8, "YTD MARGIN £", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold("I" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_Cell(RowNumber, 9, "YTD MARGIN %", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Bold("J" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":J" + (RowNumber + 1), Colour, Color.Black, SheetNumber);

               RowNumber++;

               List<string> newBusinessList = new List<string>();
               List<string> notinBudgetList = new List<string>();
               foreach (DataRow Row in YTDSales.Rows)
               {
                  BudgetModel currentRow = BudgetList.Where(w => w.ExistingCustomers == Classes.Global.ConvertToString(Row["Name"])).FirstOrDefault();
                  if (currentRow != null)
                  {
                     if (currentRow.Section != "NEW BUSINESS IN " + Year)
                     {
                        if (currentRow.ExistingCustomers != "")
                        {
                           sSheet.Set_Cell(RowNumber, 0, Row["Name"], SheetNumber);
                           sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,3,FALSE), 0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O," + (MonthColumnIndex) + ",FALSE), 0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 3, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,3,FALSE), 0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O,14,FALSE), 0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,4,FALSE), 0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 7, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,6,FALSE), 0)", SheetNumber, "#,##0.0");
                           sSheet.Set_Formula(RowNumber, 8, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,4,FALSE), 0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,6,FALSE), 0)", SheetNumber, "#,##0.0");

                           RowNumber++;
                        }
                     }
                     else
                        newBusinessList.Add(currentRow.ExistingCustomers);
                  }
                  else
                     notinBudgetList.Add(Classes.Global.ConvertToString(Row["Name"]));
               }

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "New for " + Year + "", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Bold("A" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;
               int OtherRowNumber = RowNumber;

               foreach (string newBusiness in newBusinessList)
               {
                  DataRow[] businessRows = YTDSales.Select("Name = '" + newBusiness + "'");
                  if (businessRows.Length > 0)
                  {
                     sSheet.Set_Cell(RowNumber, 0, businessRows[0]["Name"], SheetNumber);
                     sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,3,FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O," + (MonthColumnIndex) + ",FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 3, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,3,FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O,14,FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,4,FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 7, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,6,FALSE), 0)", SheetNumber, "#,##0.0");
                     sSheet.Set_Formula(RowNumber, 8, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,4,FALSE), 0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,6,FALSE), 0)", SheetNumber, "#,##0.0");

                     RowNumber++;
                  }
               }

               sSheet.Set_FontColour("A2:A" + (RowNumber + 1), LightGreen, Color.Black, SheetNumber);
               sSheet.Set_FontColour("A" + (OtherRowNumber - 2) + ":J" + OtherRowNumber, Color.LightGray, Color.Black, SheetNumber);
               sSheet.Set_FontColour("F2:F" + OtherRowNumber, Color.LightGray, Color.Black, SheetNumber);

               sSheet.Set_Formula(RowNumber, 1, "=SUM(B2:B" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(C2:C" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=SUM(D2:D" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=SUM(E2:E" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUM(G2:G" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "=AVERAGE(H2:H" + (RowNumber) + ")", SheetNumber, "#,##0.0");
               sSheet.Set_Formula(RowNumber, 8, "=SUM(I2:I" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "=AVERAGE(J2:J" + (RowNumber) + ")", SheetNumber, "#,##0.0");

               sSheet.Set_Bold_Range("B" + (RowNumber + 1) + ":J" + (RowNumber + 1), true, SheetNumber);

               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":J" + (RowNumber + 1), Colour, Color.Black, SheetNumber);

               int TotalIndex = RowNumber;

               RowNumber += 2;

               sSheet.Set_FontColour("A" + (RowNumber) + ":J" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, "OTHERS NO BUDGET " + Year + "", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Bold("A" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;
               OtherRowNumber = RowNumber;

               foreach (string noBudget in notinBudgetList)
               {
                  bool Outdated = OutdatedList.Contains(Classes.Global.ConvertToString(noBudget).Trim());
                  if (!Outdated)
                  {
                     DataRow[] noBudgetRows = YTDSales.Select("Name = '" + noBudget + "'");
                     if (noBudgetRows.Length > 0)
                     {
                        sSheet.Set_Cell(RowNumber, 0, noBudgetRows[0]["Name"], SheetNumber);
                        sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,3,FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O," + (MonthColumnIndex) + ",FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 3, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,3,FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O,14,FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,4,FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 7, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:O,6,FALSE), 0)", SheetNumber, "#,##0.0");
                        sSheet.Set_Formula(RowNumber, 8, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,4,FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'YTD SALES'!A:O,6,FALSE), 0)", SheetNumber, "#,##0.0");

                        RowNumber++;
                     }
                  }
               }

               sSheet.Set_FontColour("A" + (OtherRowNumber + 1) + ":A" + (RowNumber), LightGreen, Color.Black, SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, "TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Formula(RowNumber, 1, "=SUM(B" + (OtherRowNumber + 1) + ":B" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(C" + (OtherRowNumber + 1) + ":C" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=SUM(D" + (OtherRowNumber + 1) + ":D" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=SUM(E" + (OtherRowNumber + 1) + ":E" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUM(G" + (OtherRowNumber + 1) + ":G" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "=AVERAGE(H" + (OtherRowNumber + 1) + ":H" + (RowNumber) + ")", SheetNumber, "#,##0.0");
               sSheet.Set_Formula(RowNumber, 8, "=SUM(I" + (OtherRowNumber + 1) + ":I" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "=AVERAGE(J" + (OtherRowNumber + 1) + ":J" + (RowNumber) + ")", SheetNumber, "#,##0.0");

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":J" + (RowNumber + 1), true, SheetNumber);

               int OthersTotalIndex = RowNumber;

               RowNumber++;

               sSheet.Set_AllBorders("A1:J" + RowNumber, Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_OutsideBorders("A1:J" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "GRAND TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Formula(RowNumber, 1, "=B" + (TotalIndex + 1) + "+B" + (OthersTotalIndex + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (TotalIndex + 1) + "+C" + (OthersTotalIndex + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=D" + (TotalIndex + 1) + "+D" + (OthersTotalIndex + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=E" + (TotalIndex + 1) + "+E" + (OthersTotalIndex + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=G" + (TotalIndex + 1) + "+G" + (OthersTotalIndex + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "=I" + (TotalIndex + 1) + "+I" + (OthersTotalIndex + 1), SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "TOTAL FROM OTHER SHEET", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Formula(RowNumber, 1, "='CURRENT MONTH TURNOVER SUMMARY'!C" + MonthlyTotalRowNumber, SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + Year + " BUDGET'!" + sSheet.GetExcelColumnName(MonthColumnIndex) + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='YTD SALES'!C" + YTDTotalRowNumber, SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + Year + " BUDGET'!N" + BudgetTotalRow, SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='CURRENT MONTH TURNOVER SUMMARY'!D" + MonthlyTotalRowNumber, SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='YTD SALES'!D" + YTDTotalRowNumber, SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "VARIANCE", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Formula(RowNumber, 1, "=B" + (RowNumber) + "-B" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (RowNumber) + "-C" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=D" + (RowNumber) + "-D" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=E" + (RowNumber) + "-E" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=G" + (RowNumber) + "-G" + (RowNumber - 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "=I" + (RowNumber) + "-I" + (RowNumber - 1), SheetNumber, "£ #,##0");

               sSheet.Set_OutsideBorders("A" + (RowNumber - 1) + ":J" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Auto_fit(0, 13, SheetNumber);

               /**************************************************************************************************************************
               * MONTH SALES PER CUSTOMER SHEETS
               **************************************************************************************************************************/

               SheetNumber++;
               EndDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddMonths(1).AddSeconds(-1).ToString("yyyy-MM-dd");
               List<PriorYearSalesModel> ThisYearSales = MonthSalesPerCustomerSheets(sSheet, SheetNumber, Year, EndDate, LightGreen);
               ThisYearMonthSalesPerCustomer = SheetNumber;
               int ThisYearRowCount = sSheet.GetWorksheetRange(Year + " MONTH SALES PER CUSTOMER").RowCount;

               SheetNumber++;
               EndDate = LastYear + "-12-31";
               List<PriorYearSalesModel> LastYearSales = MonthSalesPerCustomerSheets(sSheet, SheetNumber, LastYear, EndDate, LightGreen, MonthNo);
               LastYearMonthSalesPerCustomer = SheetNumber;
               int LastYearRowCount = sSheet.GetWorksheetRange(LastYear + " MONTH SALES PER CUSTOMER").RowCount;

               SheetNumber++;
               LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-2).ToString("yyyy");
               EndDate = LastYear + "-12-31";
               List<PriorYearSalesModel> TwoYearsSales = MonthSalesPerCustomerSheets(sSheet, SheetNumber, LastYear, EndDate, LightGreen);
               PriorYearMonthSalesPerCustomer = SheetNumber;
               int Last2YearRowCount = sSheet.GetWorksheetRange(LastYear + " MONTH SALES PER CUSTOMER").RowCount;

               // Resetting afterwards
               LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy");
               EndDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddMonths(1).AddSeconds(-1).ToString("yyyy-MM-dd");

               /**************************************************************************************************************************
               * NEW B V BUDGET
               *************************************************************************************************************************/
               Color DarkTeal = System.Drawing.ColorTranslator.FromHtml("#D9E1F2");
               Color GreenAccent = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");

               SheetNumber++;

               sSheet.Insert_Worksheet("2.NEW B V BUDGET", SheetNumber);
               NewBvBudget = SheetNumber;
               RowNumber = 0;

               List<string> NameList = new List<string>();

               for (int i = NewBusinessWonStart; i < NewBusinessWonEnd; i++)
               {
                  var Value = sSheet.Get_Cell_Text(i, 0, Year + " BUDGET");
                  if (Value != null && Value != "")
                     NameList.Add(Value);
               }

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS WON IN " + LastYear, SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 1, "JAN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "FEB", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "MAR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "APR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "MAY", SheetNumber);
               sSheet.Set_Cell(RowNumber, 6, "JUN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 7, "JUL", SheetNumber);
               sSheet.Set_Cell(RowNumber, 8, "AUG", SheetNumber);
               sSheet.Set_Cell(RowNumber, 9, "SEP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 10, "OCT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 11, "NOV", SheetNumber);
               sSheet.Set_Cell(RowNumber, 12, "DEC", SheetNumber);
               sSheet.Set_Cell(RowNumber, 13, "", SheetNumber);
               sSheet.Set_Cell(RowNumber, 14, "YTD", SheetNumber);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":O" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":O" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               RowNumber++;
               int newBusinessWonBudget = RowNumber;
               int MonthlyBudgetRow = 0;

               int newBusinessforloop = NewBusinessWonStart;

               foreach (string Name in NameList)
               {
                  sqlstring = "SELECT tbl_Invoice.Invoice_Number, tbl_Customer.Name, tbl_Invoice.Invoice_Date " +
                      "FROM tbl_Invoice INNER JOIN " +
                      "tbl_Customer ON tbl_Invoice.CustomerID = tbl_Customer.CustomerID " +
                      "WHERE(tbl_Customer.Name = N'" + Name + "') " +
                      "ORDER BY tbl_Invoice.Invoice_Date ";

                  DataTable newBdt = Invoices.RetrieveDataTable(sqlstring);

                  var MonthStart = 1;
                  if (newBdt.Rows.Count > 0)
                     MonthStart = Classes.Global.ConvertToDateTime(newBdt.Rows[0]["Invoice_Date"]).Month;
                  int StartColumn = 1;

                  sSheet.Set_Cell(RowNumber, 0, Name, SheetNumber);

                  for (int i = 0; i < MonthStart; i++)
                  {
                     sSheet.Set_Formula(RowNumber, StartColumn, "='" + Year + " BUDGET'!" + sSheet.GetExcelColumnName(StartColumn + 1) + (newBusinessforloop + 1), SheetNumber, "£ #,##0");

                     StartColumn++;
                  }

                  sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 14, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

                  newBusinessforloop++;
                  RowNumber++;

                  //for (int i = NewBusinessWonStart; i < NewBusinessWonEnd; i++)
                  //{
                  //    sSheet.Set_Formula(RowNumber, 0, "='" + Year + " BUDGET'!A" + (i + 1), SheetNumber);
                  //    sSheet.Set_Formula(RowNumber, 1, "='" + Year + " BUDGET'!B" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 2, "='" + Year + " BUDGET'!C" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 3, "='" + Year + " BUDGET'!D" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 4, "='" + Year + " BUDGET'!E" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 5, "='" + Year + " BUDGET'!F" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 6, "='" + Year + " BUDGET'!G" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 7, "='" + Year + " BUDGET'!H" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 8, "='" + Year + " BUDGET'!I" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 9, "='" + Year + " BUDGET'!J" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 10, "='" + Year + " BUDGET'!K" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 11, "='" + Year + " BUDGET'!L" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 12, "='" + Year + " BUDGET'!M" + (i + 1), SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                  //    sSheet.Set_Formula(RowNumber, 14, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

                  //    RowNumber++;
                  //}
               }

               sSheet.Set_Cell(RowNumber, 0, "MONTHLY BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=SUM(B2:B" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(C2:C" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=SUM(D2:D" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=SUM(E2:E" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "=SUM(F2:F" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUM(G2:G" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "=SUM(H2:H" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "=SUM(I2:I" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "=SUM(J2:J" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=SUM(K2:K" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=SUM(L2:L" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=SUM(M2:M" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(N2:N" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 14, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               MonthlyBudgetRow = RowNumber;

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":O" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (newBusinessWonBudget + 1) + ":O" + (RowNumber + 1), DarkTeal, Color.Black, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS IN " + Year, SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 1, "JAN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "FEB", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "MAR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "APR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "MAY", SheetNumber);
               sSheet.Set_Cell(RowNumber, 6, "JUN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 7, "JUL", SheetNumber);
               sSheet.Set_Cell(RowNumber, 8, "AUG", SheetNumber);
               sSheet.Set_Cell(RowNumber, 9, "SEP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 10, "OCT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 11, "NOV", SheetNumber);
               sSheet.Set_Cell(RowNumber, 12, "DEC", SheetNumber);
               sSheet.Set_Cell(RowNumber, 13, "", SheetNumber);
               sSheet.Set_Cell(RowNumber, 14, "YTD", SheetNumber);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":O" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":O" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               RowNumber++;
               int newBusinessthisSheet = RowNumber;

               if (newBusinessEnd == 0)
                  newBusinessEnd = BudgetTotalRow - 2;

               for (int i = newBusinessStart; i < newBusinessEnd; i++)
               {
                  sSheet.Set_Formula(RowNumber, 0, "='" + Year + " BUDGET'!A" + (i + 1), SheetNumber);
                  sSheet.Set_Formula(RowNumber, 1, "='" + Year + " BUDGET'!B" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 2, "='" + Year + " BUDGET'!C" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 3, "='" + Year + " BUDGET'!D" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 4, "='" + Year + " BUDGET'!E" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 5, "='" + Year + " BUDGET'!F" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 6, "='" + Year + " BUDGET'!G" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 7, "='" + Year + " BUDGET'!H" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 8, "='" + Year + " BUDGET'!I" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 9, "='" + Year + " BUDGET'!J" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 10, "='" + Year + " BUDGET'!K" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 11, "='" + Year + " BUDGET'!L" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 12, "='" + Year + " BUDGET'!M" + (i + 1), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 14, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 0, "MONTHLY BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=SUM(B" + (newBusinessthisSheet + 1) + ":B" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(C" + (newBusinessthisSheet + 1) + ":C" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=SUM(D" + (newBusinessthisSheet + 1) + ":D" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=SUM(E" + (newBusinessthisSheet + 1) + ":E" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "=SUM(F" + (newBusinessthisSheet + 1) + ":F" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUM(G" + (newBusinessthisSheet + 1) + ":G" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "=SUM(H" + (newBusinessthisSheet + 1) + ":H" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "=SUM(I" + (newBusinessthisSheet + 1) + ":I" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "=SUM(J" + (newBusinessthisSheet + 1) + ":J" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=SUM(K" + (newBusinessthisSheet + 1) + ":K" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=SUM(L" + (newBusinessthisSheet + 1) + ":L" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=SUM(M" + (newBusinessthisSheet + 1) + ":M" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(N" + (newBusinessthisSheet + 1) + ":N" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 14, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               int newBusinessBudgetEnd = RowNumber;

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":O" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (newBusinessthisSheet + 1) + ":O" + (RowNumber + 1), GreenAccent, Color.Black, SheetNumber);

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS WON IN " + LastYear + " IMPACTING " + Year, SheetNumber);
               sSheet.Set_Cell(RowNumber, 1, "JAN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "FEB", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "MAR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "APR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "MAY", SheetNumber);
               sSheet.Set_Cell(RowNumber, 6, "JUN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 7, "JUL", SheetNumber);
               sSheet.Set_Cell(RowNumber, 8, "AUG", SheetNumber);
               sSheet.Set_Cell(RowNumber, 9, "SEP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 10, "OCT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 11, "NOV", SheetNumber);
               sSheet.Set_Cell(RowNumber, 12, "DEC", SheetNumber);
               sSheet.Set_Cell(RowNumber, 13, "YTD", SheetNumber);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               RowNumber++;
               int newBusinessWonSales = RowNumber;

               for (int i = NewBusinessWonStart; i < NewBusinessWonEnd; i++)
               {
                  sSheet.Set_Formula(RowNumber, 0, "='" + Year + " BUDGET'!A" + (i + 1), SheetNumber);
                  sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,2,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,5,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 3, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,8,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,11,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,14,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,17,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 7, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,20,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 8, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,23,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,26,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 10, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,29,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 11, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,32,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 12, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,35,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 0, "MONTHLY SALES", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=IFERROR(SUM(B" + (newBusinessWonSales + 1) + ":B" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=IFERROR(SUM(C" + (newBusinessWonSales + 1) + ":C" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(SUM(D" + (newBusinessWonSales + 1) + ":D" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=IFERROR(SUM(E" + (newBusinessWonSales + 1) + ":E" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "=IFERROR(SUM(F" + (newBusinessWonSales + 1) + ":F" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=IFERROR(SUM(G" + (newBusinessWonSales + 1) + ":G" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(SUM(H" + (newBusinessWonSales + 1) + ":H" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "=IFERROR(SUM(I" + (newBusinessWonSales + 1) + ":I" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "=IFERROR(SUM(J" + (newBusinessWonSales + 1) + ":J" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=IFERROR(SUM(K" + (newBusinessWonSales + 1) + ":K" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=IFERROR(SUM(L" + (newBusinessWonSales + 1) + ":L" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=IFERROR(SUM(M" + (newBusinessWonSales + 1) + ":M" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               int MonthlySalesRow = RowNumber;

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;

               sSheet.Set_Formula(RowNumber, 1, "=B" + (RowNumber) + "/B" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (RowNumber) + "/C" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 3, "=D" + (RowNumber) + "/D" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=E" + (RowNumber) + "/E" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=F" + (RowNumber) + "/F" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 6, "=G" + (RowNumber) + "/G" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 7, "=H" + (RowNumber) + "/H" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 8, "=I" + (RowNumber) + "/I" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 9, "=J" + (RowNumber) + "/J" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 10, "=K" + (RowNumber) + "/K" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 11, "=L" + (RowNumber) + "/L" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 12, "=M" + (RowNumber) + "/M" + (newBusinessthisSheet - 1), SheetNumber, "#,##0 %");

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (newBusinessWonSales + 1) + ":N" + (RowNumber + 1), DarkTeal, Color.Black, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS IN " + Year, SheetNumber);
               sSheet.Set_Cell(RowNumber, 1, "JAN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "FEB", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "MAR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "APR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "MAY", SheetNumber);
               sSheet.Set_Cell(RowNumber, 6, "JUN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 7, "JUL", SheetNumber);
               sSheet.Set_Cell(RowNumber, 8, "AUG", SheetNumber);
               sSheet.Set_Cell(RowNumber, 9, "SEP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 10, "OCT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 11, "NOV", SheetNumber);
               sSheet.Set_Cell(RowNumber, 12, "DEC", SheetNumber);
               sSheet.Set_Cell(RowNumber, 13, "YTD", SheetNumber);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               RowNumber++;
               int newBusinessSales = RowNumber;

               for (int i = newBusinessStart; i < newBusinessEnd; i++)
               {
                  sSheet.Set_Formula(RowNumber, 0, "='" + Year + " BUDGET'!A" + (i + 1), SheetNumber);
                  sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,2,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,5,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 3, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,8,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,11,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,14,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,17,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 7, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,20,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 8, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,23,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,26,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 10, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,29,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 11, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,32,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 12, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,35,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

                  RowNumber++;
               }

               int MonthlySaleswBudget = RowNumber;

               sSheet.Set_Cell(RowNumber, 0, "MONTHLY SALES - 'NEW IN " + Year + "' WITH BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=IFERROR(SUM(B" + (newBusinessSales + 1) + ":B" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=IFERROR(SUM(C" + (newBusinessSales + 1) + ":C" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(SUM(D" + (newBusinessSales + 1) + ":D" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=IFERROR(SUM(E" + (newBusinessSales + 1) + ":E" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "=IFERROR(SUM(F" + (newBusinessSales + 1) + ":F" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=IFERROR(SUM(G" + (newBusinessSales + 1) + ":G" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(SUM(H" + (newBusinessSales + 1) + ":H" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "=IFERROR(SUM(I" + (newBusinessSales + 1) + ":I" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "=IFERROR(SUM(J" + (newBusinessSales + 1) + ":J" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=IFERROR(SUM(K" + (newBusinessSales + 1) + ":K" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=IFERROR(SUM(L" + (newBusinessSales + 1) + ":L" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=IFERROR(SUM(M" + (newBusinessSales + 1) + ":M" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;

               sSheet.Set_Formula(RowNumber, 1, "=B" + (RowNumber) + "/B" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (RowNumber) + "/C" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 3, "=D" + (RowNumber) + "/D" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=E" + (RowNumber) + "/E" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=F" + (RowNumber) + "/F" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 6, "=G" + (RowNumber) + "/G" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 7, "=H" + (RowNumber) + "/H" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 8, "=I" + (RowNumber) + "/I" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 9, "=J" + (RowNumber) + "/J" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 10, "=K" + (RowNumber) + "/K" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 11, "=L" + (RowNumber) + "/L" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 12, "=M" + (RowNumber) + "/M" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");

               sSheet.Set_FontColour("A" + (newBusinessSales + 1) + ":N" + (RowNumber + 1), GreenAccent, Color.Black, SheetNumber);
               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS NO BUDGET", SheetNumber);
               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;
               int newBusinessnoBudget = RowNumber;

               for (int i = 0; i < newCustomerNoBudgetTotalRows; i++)
               {
                  sSheet.Set_Formula(RowNumber, 0, "='NEW BUSINESS NO BUDGET'!A" + (i + 1), SheetNumber);
                  sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,2,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,5,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 3, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,8,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,11,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,14,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,17,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 7, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,20,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 8, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,23,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,26,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 10, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,29,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 11, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,32,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 12, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AN,35,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

                  RowNumber++;
               }

               int totalNewNoBudgetRow = RowNumber;

               sSheet.Set_Cell(RowNumber, 0, "TOTAL NEW NO BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=IFERROR(SUM(B" + (newBusinessnoBudget + 1) + ":B" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=IFERROR(SUM(C" + (newBusinessnoBudget + 1) + ":C" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(SUM(D" + (newBusinessnoBudget + 1) + ":D" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=IFERROR(SUM(E" + (newBusinessnoBudget + 1) + ":E" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "=IFERROR(SUM(F" + (newBusinessnoBudget + 1) + ":F" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=IFERROR(SUM(G" + (newBusinessnoBudget + 1) + ":G" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(SUM(H" + (newBusinessnoBudget + 1) + ":H" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "=IFERROR(SUM(I" + (newBusinessnoBudget + 1) + ":I" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "=IFERROR(SUM(J" + (newBusinessnoBudget + 1) + ":J" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=IFERROR(SUM(K" + (newBusinessnoBudget + 1) + ":K" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=IFERROR(SUM(L" + (newBusinessnoBudget + 1) + ":L" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=IFERROR(SUM(M" + (newBusinessnoBudget + 1) + ":M" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               RowNumber++;

               int totalVBudgetRow = RowNumber;

               sSheet.Set_Cell(RowNumber, 0, "TOTAL V BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=B" + (MonthlySaleswBudget + 1) + "+B" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (MonthlySaleswBudget + 1) + "+C" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=D" + (MonthlySaleswBudget + 1) + "+D" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "=E" + (MonthlySaleswBudget + 1) + "+E" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "=F" + (MonthlySaleswBudget + 1) + "+F" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=G" + (MonthlySaleswBudget + 1) + "+G" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "=H" + (MonthlySaleswBudget + 1) + "+H" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "=I" + (MonthlySaleswBudget + 1) + "+I" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "=J" + (MonthlySaleswBudget + 1) + "+J" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=K" + (MonthlySaleswBudget + 1) + "+K" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=L" + (MonthlySaleswBudget + 1) + "+L" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "=M" + (MonthlySaleswBudget + 1) + "+M" + (RowNumber), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=N" + (MonthlySaleswBudget + 1) + "+N" + (RowNumber), SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "VARIANCE AGAINST BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=B" + (RowNumber) + "/B" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (RowNumber) + "/C" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 3, "=D" + (RowNumber) + "/D" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=E" + (RowNumber) + "/E" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=F" + (RowNumber) + "/F" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 6, "=G" + (RowNumber) + "/G" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 7, "=H" + (RowNumber) + "/H" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 8, "=I" + (RowNumber) + "/I" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 9, "=J" + (RowNumber) + "/J" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 10, "=K" + (RowNumber) + "/K" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 11, "=L" + (RowNumber) + "/L" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 12, "=M" + (RowNumber) + "/M" + (newBusinessBudgetEnd + 1), SheetNumber, "#,##0 %");

               sSheet.Set_Bold_Range("A" + (RowNumber - 1) + ":N" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (newBusinessnoBudget) + ":N" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               // New Dean Items on New B v Budget
               // Monthly Sales

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "MTHLY SALES - BUD 'NEW IN " + Year + "'", SheetNumber);
               for (int i = 1; i < 13; i++)
                  sSheet.Set_Formula(RowNumber, i, "=" + sSheet.GetExcelColumnName(i + 1) + (newBusinessnoBudget - 2), SheetNumber, "£ #,##0");

               // Total new no Budget

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "TOTAL NEW NO BUDGET", SheetNumber);
               for (int i = 1; i < 13; i++)
                  sSheet.Set_Formula(RowNumber, i, "=" + sSheet.GetExcelColumnName(i + 1) + (totalNewNoBudgetRow + 1), SheetNumber, "£ #,##0");

               // New Business

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS IN " + Year, SheetNumber);
               for (int i = 1; i < 13; i++)
                  sSheet.Set_Formula(RowNumber, i, "=" + sSheet.GetExcelColumnName(i + 1) + (newBusinessBudgetEnd + 1), SheetNumber, "£ #,##0");

               // New Business FY

               RowNumber += 2;

               int NewBusinessFYStart = RowNumber;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS IN " + Year, SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "=SUM(B" + (RowNumber - 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber - 1) + ")", SheetNumber, "£ #,##0");

               // Monthly Sales FY

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "MNTHLY SALES - BUD 'NEW IN " + Year + "'", SheetNumber);
               sSheet.Set_Formula(RowNumber, 2, "=SUM(B" + (RowNumber - 5) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber - 5) + ")", SheetNumber, "£ #,##0");

               // No Budget FY

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "MNTHLY SALES - 'NEW NO BUD'" + Year, SheetNumber);
               sSheet.Set_Formula(RowNumber, 2, "=SUM(B" + (RowNumber - 5) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber - 5) + ")", SheetNumber, "£ #,##0");

               // New Business Won Smaller Table

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS WON IN " + LastYear + " IMPACTING " + Year, SheetNumber);
               sSheet.Set_Cell(RowNumber, 1, "BUD YTD " + Month, SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "ACT YTD", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "BUD FY", SheetNumber);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               RowNumber++;
               int c = newBusinessWonBudget + 1;
               int oc = newBusinessWonSales + 1;
               for (int i = NewBusinessWonStart; i < NewBusinessWonEnd; i++)
               {
                  sSheet.Set_Formula(RowNumber, 0, "=A" + (c), SheetNumber);
                  sSheet.Set_Formula(RowNumber, 1, "=O" + (c), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 2, "=N" + (oc), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 3, "=N" + (c), SheetNumber, "£ #,##0");

                  sSheet.Set_FontColour("A" + (RowNumber + 1), DarkTeal, Color.Black, SheetNumber);
                  sSheet.Set_AllBorders("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
                  c++;
                  oc++;
                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 0, "MONTHLY BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=O" + (c), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=N" + (oc), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=N" + (c), SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Formula(RowNumber, 2, "=B" + (RowNumber) + "-C" + (RowNumber), SheetNumber, "£ #,##0");

               // New Business Smaller Table

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS IN " + Year, SheetNumber);
               sSheet.Set_Cell(RowNumber, 1, "BUD YTD " + Month, SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "ACT YTD", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "BUD FY", SheetNumber);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               RowNumber++;
               c = newBusinessthisSheet + 1;
               oc = newBusinessSales + 1;
               for (int i = newBusinessStart; i < newBusinessEnd; i++)
               {
                  sSheet.Set_Formula(RowNumber, 0, "=A" + (c), SheetNumber);
                  sSheet.Set_Formula(RowNumber, 1, "=O" + (c), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 2, "=N" + (oc), SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 3, "=N" + (c), SheetNumber, "£ #,##0");

                  sSheet.Set_FontColour("A" + (RowNumber + 1), GreenAccent, Color.Black, SheetNumber);
                  sSheet.Set_AllBorders("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
                  c++;
                  oc++;
                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 0, "MNTHLY SALES - BUD 'NEW IN " + Year + "'", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=O" + (c), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=N" + (oc), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=N" + (c), SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Formula(RowNumber, 2, "=B" + (RowNumber) + "-C" + (RowNumber), SheetNumber, "£ #,##0");

               // New Business No Budget Smaller Table

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "NEW BUS NO BUDGET", SheetNumber);
               sSheet.Set_Cell(RowNumber, 1, "ACTUAL YTD", SheetNumber);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":B" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_FontColour("A" + (RowNumber + 1) + ":B" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

               RowNumber++;

               c = newBusinessnoBudget + 1;
               for (int i = 0; i < newCustomerNoBudgetTotalRows; i++)
               {
                  sSheet.Set_Formula(RowNumber, 0, "='NEW BUSINESS NO BUDGET'!A" + (i + 1), SheetNumber);
                  sSheet.Set_Formula(RowNumber, 1, "=N" + (c), SheetNumber, "£ #,##0");
                  sSheet.Set_FontColour("A" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);
                  sSheet.Set_AllBorders("A" + (RowNumber + 1) + ":B" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);

                  c++;
                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 0, "TOTAL NEW NO BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=N" + (c), SheetNumber, "£ #,##0");

               // New Business Won
               sms.Set_Chart(sSheet, "B" + (MonthlyBudgetRow + 1) + ":M" + (MonthlyBudgetRow + 1), "B" + (MonthlySalesRow + 1) + ":M" + (MonthlySalesRow + 1), "A" + (MonthlyBudgetRow + 1), "A" + (MonthlySalesRow + 1),
               "B1:M1", "B1:M1", "Q2", "AE14", SheetNumber, ChartType.ColumnClustered, Color.DarkGray, Color.Purple, LightGreen, LightGreen, LegendPosition.Bottom, "NEW BUSINESS WON IN " + LastYear + " IMPACTING " + Year,
               null, null, null, null, null, null, false, false, true);

               // Budgeted New Business
               sms.Set_Chart(sSheet, "B" + (newBusinessBudgetEnd + 1) + ":M" + (newBusinessBudgetEnd + 1), "B" + (MonthlySaleswBudget + 1) + ":M" + (MonthlySaleswBudget + 1), "A" + (newBusinessBudgetEnd + 1), "A" + (MonthlySaleswBudget + 1),
                   "B1:M1", "B1:M1", "Q16", "AG26", SheetNumber, ChartType.ColumnClustered, Color.DarkGray, Color.Green, LightGreen, LightGreen, LegendPosition.Bottom, "BUDGETED NEW BUSINESS IN " + Year,
                   null, null, null, null, null, null, false, false, true);

               // Total New Business
               sms.Set_Chart(sSheet, "B" + (MonthlySaleswBudget + 1) + ":M" + (MonthlySaleswBudget + 1), "B" + (totalNewNoBudgetRow + 1) + ":M" + (totalNewNoBudgetRow + 1), "A" + (MonthlySaleswBudget + 1), "A" + (totalNewNoBudgetRow + 1),
                   "B1:M1", "B1:M1", "Q28", "AB44", SheetNumber, ChartType.ColumnClustered, Color.Gray, Color.Teal, LightGreen, Color.DarkGray, LegendPosition.Bottom, "TOTAL NEW BUSINESS IN " + Year + " v BUDGET",
                   "B" + (totalVBudgetRow + 1) + ":M" + (totalVBudgetRow + 1), "A" + (totalVBudgetRow + 1), "B1:M1", "B" + (newBusinessBudgetEnd + 1) + ":M" + (newBusinessBudgetEnd + 1), "A" + (newBusinessBudgetEnd + 1), "B1:M1", true);

               // New Business Budget v Monthly Sales
               sms.Set_Chart(sSheet, "B" + (NewBusinessFYStart + 1) + ":C" + (NewBusinessFYStart + 1), "B" + (NewBusinessFYStart + 2) + ":C" + (NewBusinessFYStart + 2), "A" + (NewBusinessFYStart + 1), "A" + (NewBusinessFYStart + 2),
                   "A" + (NewBusinessFYStart + 1), "A" + (NewBusinessFYStart + 2), "P70", "W85", SheetNumber, ChartType.ColumnStacked, Color.Blue, Color.Orange, Color.Green, LightGreen, LegendPosition.Bottom, "NEW BUSINESS IN " + Year + " VS MONTHLY SALES FY",
                   "B" + (NewBusinessFYStart + 3) + ":C" + (NewBusinessFYStart + 3), "A" + (NewBusinessFYStart + 3), "A" + (NewBusinessFYStart + 3), null, null, null, false, false, true);


               sSheet.Set_Column_Width(0, 41.57, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(1, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(2, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(3, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(4, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(5, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(6, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(7, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(8, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(9, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(10, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(11, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(12, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(13, 10.86, "2.NEW B V BUDGET");
               sSheet.Set_Column_Width(14, 10.86, "2.NEW B V BUDGET");

               /**************************************************************************************************************************
               * SALES V PRIOR YEARS
               *************************************************************************************************************************/

               RowNumber = 0;
               SheetNumber++;

               sSheet.Insert_Worksheet("3.SALES V PRIOR YEARS", SheetNumber);
               TotalSalesvPriorYears = SheetNumber;

               sSheet.Set_Cell(RowNumber, 0, LastYear, SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Cell(RowNumber, 1, "JANUARY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 2, "FEBRUARY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 3, "MARCH", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 4, "APRIL", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 5, "MAY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 6, "JUNE", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 7, "JULY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 8, "AUGUST", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 9, "SEPTEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 10, "OCTOBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 11, "NOVEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 12, "DECEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 13, "TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Left);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "SALES " + LastYear, SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "='" + LastYear + " MONTH SALES PER CUSTOMER'!B" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + LastYear + " MONTH SALES PER CUSTOMER'!E" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + LastYear + " MONTH SALES PER CUSTOMER'!H" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + LastYear + " MONTH SALES PER CUSTOMER'!K" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + LastYear + " MONTH SALES PER CUSTOMER'!N" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Q" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + LastYear + " MONTH SALES PER CUSTOMER'!T" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + LastYear + " MONTH SALES PER CUSTOMER'!W" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Z" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AC" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AF" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AI" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               int SalesLastYearsRow = RowNumber;

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "BUDGET", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "='" + LastYear + " BUDGET'!B" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + LastYear + " BUDGET'!C" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + LastYear + " BUDGET'!D" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + LastYear + " BUDGET'!E" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + LastYear + " BUDGET'!F" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + LastYear + " BUDGET'!G" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + LastYear + " BUDGET'!H" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + LastYear + " BUDGET'!I" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + LastYear + " BUDGET'!J" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + LastYear + " BUDGET'!K" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + LastYear + " BUDGET'!L" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + LastYear + " BUDGET'!M" + (priorYearBudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "PRIOR YEAR", SheetNumber);
               LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-2).ToString("yyyy");
               sSheet.Set_Formula(RowNumber, 1, "='" + LastYear + " MONTH SALES PER CUSTOMER'!B" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + LastYear + " MONTH SALES PER CUSTOMER'!E" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + LastYear + " MONTH SALES PER CUSTOMER'!H" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + LastYear + " MONTH SALES PER CUSTOMER'!K" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + LastYear + " MONTH SALES PER CUSTOMER'!N" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Q" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + LastYear + " MONTH SALES PER CUSTOMER'!T" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + LastYear + " MONTH SALES PER CUSTOMER'!W" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Z" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AC" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AF" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AI" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               sSheet.Set_AllBorders("A1:M4", Color.Black, BorderLineStyle.Thin, SheetNumber);

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "BUDGET %", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "=B2/B3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 2, "=C2/C3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 3, "=D2/D3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=E2/E3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=F2/F3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 6, "=G2/G3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 7, "=H2/H3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 8, "=I2/I3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 9, "=J2/J3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 10, "=K2/K3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 11, "=L2/L3", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 12, "=M2/M3", SheetNumber, "#,##0 %");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "PRIOR YR VARIANCE", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "=(B2-B4)/B4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 2, "=(C2-C4)/C4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 3, "=(D2-D4)/D4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=(E2-E4)/E4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=(F2-F4)/F4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 6, "=(G2-G4)/G4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 7, "=(H2-H4)/H4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 8, "=(I2-I4)/I4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 9, "=(J2-J4)/J4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 10, "=(K2-K4)/K4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 11, "=(L2-L4)/L4", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 12, "=(M2-M4)/M4", SheetNumber, "#,##0 %");

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, LastYear, SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=B4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 2, "=C4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 3, "=D4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 4, "=E4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 5, "=F4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 6, "=G4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 7, "=H4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 8, "=I4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 9, "=J4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 10, "=K4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 11, "=L4/N4", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 12, "=M4/N4", SheetNumber, "#,##0.00 %");

               // Resetting afterwards
               LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, LastYear, SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=B2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 2, "=C2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 3, "=D2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 4, "=E2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 5, "=F2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 6, "=G2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 7, "=H2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 8, "=I2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 9, "=J2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 10, "=K2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 11, "=L2/N2", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 12, "=M2/N2", SheetNumber, "#,##0.00 %");

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 1, "Q1", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "Q2", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "Q3", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "Q4", SheetNumber);

               RowNumber++;

               sSheet.Set_Formula(RowNumber, 1, "=SUM(B9:D9)", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(E9:G9)", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 3, "=SUM(H9:J9)", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 4, "=SUM(K9:M9)", SheetNumber, "#,##0.00 %");

               RowNumber++;

               sSheet.Set_Formula(RowNumber, 1, "=SUM(B10:D10)", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(E10:G10)", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 3, "=SUM(H10:J10)", SheetNumber, "#,##0.00 %");
               sSheet.Set_Formula(RowNumber, 4, "=SUM(K10:M10)", SheetNumber, "#,##0.00 %");

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, Year, SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Cell(RowNumber, 1, "JANUARY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 2, "FEBRUARY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 3, "MARCH", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 4, "APRIL", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 5, "MAY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 6, "JUNE", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 7, "JULY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 8, "AUGUST", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 9, "SEPTEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 10, "OCTOBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 11, "NOVEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 12, "DECEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 13, "TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Left);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "SALES " + Year, SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "='" + Year + " MONTH SALES PER CUSTOMER'!B" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + Year + " MONTH SALES PER CUSTOMER'!E" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + Year + " MONTH SALES PER CUSTOMER'!H" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + Year + " MONTH SALES PER CUSTOMER'!K" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + Year + " MONTH SALES PER CUSTOMER'!N" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + Year + " MONTH SALES PER CUSTOMER'!Q" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + Year + " MONTH SALES PER CUSTOMER'!T" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + Year + " MONTH SALES PER CUSTOMER'!W" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + Year + " MONTH SALES PER CUSTOMER'!Z" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + Year + " MONTH SALES PER CUSTOMER'!AC" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + Year + " MONTH SALES PER CUSTOMER'!AF" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + Year + " MONTH SALES PER CUSTOMER'!AI" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               int SalesThisYearRow = RowNumber;

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "BUDGET", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "='" + Year + " BUDGET'!B" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + Year + " BUDGET'!C" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + Year + " BUDGET'!D" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + Year + " BUDGET'!E" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + Year + " BUDGET'!F" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + Year + " BUDGET'!G" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + Year + " BUDGET'!H" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + Year + " BUDGET'!I" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + Year + " BUDGET'!J" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + Year + " BUDGET'!K" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + Year + " BUDGET'!L" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + Year + " BUDGET'!M" + (BudgetTotalRow), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "PRIOR YEAR", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "='" + LastYear + " MONTH SALES PER CUSTOMER'!B" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + LastYear + " MONTH SALES PER CUSTOMER'!E" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + LastYear + " MONTH SALES PER CUSTOMER'!H" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + LastYear + " MONTH SALES PER CUSTOMER'!K" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + LastYear + " MONTH SALES PER CUSTOMER'!N" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Q" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + LastYear + " MONTH SALES PER CUSTOMER'!T" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + LastYear + " MONTH SALES PER CUSTOMER'!W" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Z" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AC" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AF" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AI" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               sSheet.Set_AllBorders("A16:M19", Color.Black, BorderLineStyle.Thin, SheetNumber);

               RowNumber += 2;

               sSheet.Set_Cell(RowNumber, 0, "BUDGET %", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "=B17/B18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 2, "=C17/C18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 3, "=D17/D18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=E17/E18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=F17/F18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 6, "=G17/G18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 7, "=H17/H18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 8, "=I17/I18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 9, "=J17/J18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 10, "=K17/K18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 11, "=L17/L18", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 12, "=M17/M18", SheetNumber, "#,##0 %");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "PRIOR YR VARIANCE", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "=B17/B19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 2, "=C17/C19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 3, "=D17/D19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=E17/E19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 5, "=F17/F19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 6, "=G17/G19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 7, "=H17/H19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 8, "=I17/I19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 9, "=J17/J19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 10, "=K17/K19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 11, "=L17/L19", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 12, "=M17/M19", SheetNumber, "#,##0 %");

               RowNumber += 21;

               sSheet.Set_Cell(RowNumber, 1, "JANUARY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 2, "FEBRUARY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 3, "MARCH", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 4, "APRIL", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 5, "MAY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 6, "JUNE", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 7, "JULY", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 8, "AUGUST", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 9, "SEPTEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 10, "OCTOBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 11, "NOVEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 12, "DECEMBER", SheetNumber, SpreadsheetHorizontalAlignment.Left);
               sSheet.Set_Cell(RowNumber, 13, "TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Left);

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":N" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;
               LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-2).ToString("yyyy");

               sSheet.Set_Cell(RowNumber, 0, "SALES " + LastYear, SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "='" + LastYear + " MONTH SALES PER CUSTOMER'!B" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + LastYear + " MONTH SALES PER CUSTOMER'!E" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + LastYear + " MONTH SALES PER CUSTOMER'!H" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + LastYear + " MONTH SALES PER CUSTOMER'!K" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + LastYear + " MONTH SALES PER CUSTOMER'!N" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Q" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + LastYear + " MONTH SALES PER CUSTOMER'!T" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + LastYear + " MONTH SALES PER CUSTOMER'!W" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Z" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AC" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AF" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AI" + (Last2YearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               int SalesAllYearsRow = RowNumber;

               RowNumber++;

               // Resetting afterwards
               LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy");

               sSheet.Set_Cell(RowNumber, 0, "SALES " + LastYear, SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "='" + LastYear + " MONTH SALES PER CUSTOMER'!B" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + LastYear + " MONTH SALES PER CUSTOMER'!E" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + LastYear + " MONTH SALES PER CUSTOMER'!H" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + LastYear + " MONTH SALES PER CUSTOMER'!K" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + LastYear + " MONTH SALES PER CUSTOMER'!N" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Q" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + LastYear + " MONTH SALES PER CUSTOMER'!T" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + LastYear + " MONTH SALES PER CUSTOMER'!W" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + LastYear + " MONTH SALES PER CUSTOMER'!Z" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AC" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AF" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + LastYear + " MONTH SALES PER CUSTOMER'!AI" + (LastYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "SALES " + Year, SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "='" + Year + " MONTH SALES PER CUSTOMER'!B" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "='" + Year + " MONTH SALES PER CUSTOMER'!E" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "='" + Year + " MONTH SALES PER CUSTOMER'!H" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 4, "='" + Year + " MONTH SALES PER CUSTOMER'!K" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 5, "='" + Year + " MONTH SALES PER CUSTOMER'!N" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "='" + Year + " MONTH SALES PER CUSTOMER'!Q" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 7, "='" + Year + " MONTH SALES PER CUSTOMER'!T" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 8, "='" + Year + " MONTH SALES PER CUSTOMER'!W" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 9, "='" + Year + " MONTH SALES PER CUSTOMER'!Z" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 10, "='" + Year + " MONTH SALES PER CUSTOMER'!AC" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 11, "='" + Year + " MONTH SALES PER CUSTOMER'!AF" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 12, "='" + Year + " MONTH SALES PER CUSTOMER'!AI" + (ThisYearRowCount), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");

               sSheet.Set_AllBorders("A43:M46", Color.Black, BorderLineStyle.Thin, SheetNumber);

               // Last Year -1 v Last Year v Year
               sms.Set_Chart(sSheet, "B" + (SalesAllYearsRow + 1) + ":M" + (SalesAllYearsRow + 1), "B" + (SalesAllYearsRow + 2) + ":M" + (SalesAllYearsRow + 2), "A" + (SalesAllYearsRow + 1), "A" + (SalesAllYearsRow + 2),
                   "B1:M1", "B1:M1", "C25", "J40", SheetNumber, ChartType.ColumnClustered, LightGreen, Color.Teal, Color.Purple, Color.Black, LegendPosition.Bottom,
                   Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-2).ToString("yyyy") + " v " + LastYear + " v " + Year, "B" + (SalesAllYearsRow + 3) + ":M" + (SalesAllYearsRow + 3),
                   "A" + (SalesAllYearsRow + 3), "B1:M1");

               //Last Year
               sms.Set_Chart(sSheet, "B" + (SalesLastYearsRow + 1) + ":M" + (SalesLastYearsRow + 1), "B" + (SalesLastYearsRow + 3) + ":M" + (SalesLastYearsRow + 3), "A" + (SalesLastYearsRow + 1), "A" + (SalesLastYearsRow + 3),
                   "B1:M1", "B1:M1", "O1", "AB16", SheetNumber, ChartType.ColumnClustered, Color.Teal, Color.Gray, LightGreen, Color.Black, LegendPosition.Bottom, LastYear + " SALES V BUDGET V PRIOR YEAR",
                   "B" + (SalesLastYearsRow + 2) + ":M" + (SalesLastYearsRow + 2), "A" + (SalesLastYearsRow + 2), "B1:M1", null, null, null, true);

               //Year
               sms.Set_Chart(sSheet, "B" + (SalesThisYearRow + 1) + ":M" + (SalesThisYearRow + 1), "B" + (SalesThisYearRow + 3) + ":M" + (SalesThisYearRow + 3), "A" + (SalesThisYearRow + 1), "A" + (SalesThisYearRow + 3),
                   "B1:M1", "B1:M1", "O18", "AB37", SheetNumber, ChartType.ColumnClustered, Color.Teal, Color.Purple, LightGreen, Color.Black, LegendPosition.Bottom, Year + " SALES V BUDGET V PRIOR YEAR",
                   "B" + (SalesThisYearRow + 2) + ":M" + (SalesThisYearRow + 2), "A" + (SalesThisYearRow + 2), "B1:M1", null, null, null, true);

               sSheet.Set_Bold_Range("A1:A50", true, SheetNumber);
               sSheet.Set_Column_Width(0, 18.29, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(1, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(2, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(3, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(4, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(5, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(6, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(7, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(8, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(9, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(10, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(11, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(12, 12, "3.SALES V PRIOR YEARS");
               sSheet.Set_Column_Width(13, 12, "3.SALES V PRIOR YEARS");

               /**************************************************************************************************************************
               * CUSTOMERS NOT BOUGHT THIS MONTH
               *************************************************************************************************************************/

               SheetNumber++;

               sSheet.Insert_Worksheet("CUSTOMERS NOT BOUGHT THIS MONTH", SheetNumber);
               CustomersNotBoughtThisMonth = SheetNumber;

               RowNumber = 0;

               sSheet.Set_Cell(RowNumber, 1, "CUSTOMER NAME", SheetNumber);

               sSheet.Set_Cell(RowNumber, 3, "CUSTOMERS NOT BOUGHT THIS MONTH THAT HAVE PURCHASED THIS YEAR", SheetNumber);

               RowNumber++;

               DataTable NotBoughtthisMonth = new DataTable();

               for (int i = 3; i < ThisYearRowCount; i++)
               {
                  var Value = sSheet.Get_Cell_Text((i - 1), (((MonthColumnIndex - 1) * 3) - 2), Year + " MONTH SALES PER CUSTOMER");
                  if (Value == null || Value == "")
                  {
                     var Name = sSheet.Get_Cell_Text((i - 1), 0, Year + " MONTH SALES PER CUSTOMER");
                     if (Name != null && Name != "")
                     {
                        sqlstring = "SELECT tbl_Invoice.Invoice_Date " +
                        "FROM tbl_Customer INNER JOIN " +
                         "tbl_Invoice ON tbl_Customer.CustomerID = tbl_Invoice.CustomerID " +
                        "WHERE(tbl_Invoice.Invoice_Date BETWEEN '" + Year + "-01-01' AND '" + EndDate + "') AND(tbl_Customer.Name = N'" + Name + "') " +
                        "ORDER BY tbl_Invoice.Invoice_Date ";

                        NotBoughtthisMonth = Invoices.RetrieveDataTable(sqlstring);

                        if (NotBoughtthisMonth.Rows.Count > 0)
                        {
                           bool hasBoughtThisMonth = false;
                           for (int j = 0; j < NotBoughtthisMonth.Rows.Count; j++)
                           {
                              if (Classes.Global.ConvertToDateTime(NotBoughtthisMonth.Rows[j]["Invoice_Date"]).Month.ToString("MMMM") == Month)
                                 hasBoughtThisMonth = true;
                           }

                           if (!hasBoughtThisMonth)
                           {
                              sSheet.Set_Formula(RowNumber, 1, "='" + Year + " MONTH SALES PER CUSTOMER'!A" + i, SheetNumber);
                              RowNumber++;
                           }
                        }
                        else
                           OutdatedList.Add(Name);

                     }
                  }
               }

               sSheet.Set_AllBorders("B1:B" + (RowNumber), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_Bold_Range("B1:D1", true, SheetNumber);
               sSheet.Set_FontColour("B1", Color.LightGray, Color.Black, SheetNumber);
               sSheet.Set_Column_Width(1, 41.57, "CUSTOMERS NOT BOUGHT THIS MONTH");
               sSheet.Set_Column_Width(3, 41.57, "CUSTOMERS NOT BOUGHT THIS MONTH");

               /**************************************************************************************************************************
               * THIS YR V LAST YR
               *************************************************************************************************************************/

               Color Pink = ColorTranslator.FromHtml("#FFCCFF");
               SheetNumber++;
               RowNumber = 0;

               sSheet.Insert_Worksheet("4.CUSTOMER £ v BDG v PY", SheetNumber);
               ThisYearvsLastYear = SheetNumber;

               sSheet.Set_Cell(RowNumber, 1, "JAN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 6, "FEB", SheetNumber);
               sSheet.Set_Cell(RowNumber, 11, "MAR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 16, "APR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 21, "MAY", SheetNumber);
               sSheet.Set_Cell(RowNumber, 26, "JUN", SheetNumber);
               sSheet.Set_Cell(RowNumber, 31, "JUL", SheetNumber);
               sSheet.Set_Cell(RowNumber, 36, "AUG", SheetNumber);
               sSheet.Set_Cell(RowNumber, 41, "SEP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 46, "OCT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 51, "NOV", SheetNumber);
               sSheet.Set_Cell(RowNumber, 56, "DEC", SheetNumber);
               sSheet.Set_Cell(RowNumber, 61, "YTD TOTAL", SheetNumber);

               sSheet.Merge_Cells("BJ1:BK1", SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "THIS YR V LAST YR", SheetNumber, SpreadsheetHorizontalAlignment.Center);

               sSheet.Set_Rotation("A2", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("A2", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Merge_Cells("A1:A2", SheetNumber);

               sSheet.Set_FontColour("A1:BK2", LightGreen, Color.Black, SheetNumber);
               sSheet.Set_FontColour("B1:BK1", Color.LightGray, Color.Black, SheetNumber);
               sSheet.Set_FontColour("BJ2:BO2", Color.LightGray, Color.Black, SheetNumber);

               for (int i = 1; i < 61; i += 5)
               {
                  sSheet.Set_Cell(RowNumber, i, LastYear, SheetNumber);
                  sSheet.Set_Cell(RowNumber, (i + 1), Year, SheetNumber);
                  sSheet.Set_Cell(RowNumber, (i + 2), "BUDGET", SheetNumber);
                  sSheet.Set_Cell(RowNumber, (i + 3), "BDG VAR", SheetNumber);
                  sSheet.Set_Cell(RowNumber, (i + 4), "PR YR VAR", SheetNumber);
               }

               sSheet.Set_Cell(RowNumber, 61, Year + " SALES", SheetNumber);
               sSheet.Set_Cell(RowNumber, 62, LastYear + " SALES", SheetNumber);
               sSheet.Set_Cell(RowNumber, 63, "BUDGET", SheetNumber);
               sSheet.Set_Cell(RowNumber, 64, "VAR AGAINST PR YR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 65, "VAR V BDG", SheetNumber);
               sSheet.Set_Cell(RowNumber, 66, "COMMENTS", SheetNumber);

               RowNumber++;

               foreach (DataRow row in YTDSales.Rows)
               {
                  bool Outdated = OutdatedList.Contains(Classes.Global.ConvertToString(row["Name"]).Trim());
                  bool inBudget = BudgetNameList.Contains(Classes.Global.ConvertToString(row["Name"]).Trim());
                  if (!Outdated || inBudget)
                  {
                     if (Classes.Global.ConvertToString(row["Name"]) != "")
                     {
                        if (Classes.Global.ConvertToString(row["Name"]) != "STANNAH STAIRLIFT EURO ACCOUNT")
                        {
                           sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]).Trim(), SheetNumber);
                           int LookupValueCurrentMonth = 2;
                           int LookupValueBudget = 2;
                           for (int i = 1; i < 61; i += 5)
                           {
                              sSheet.Set_Formula(RowNumber, i, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AK," + LookupValueCurrentMonth + ",FALSE),0)", SheetNumber, "£ #,##0.00");
                              sSheet.Set_Formula(RowNumber, i + 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AK," + LookupValueCurrentMonth + ",FALSE),0)", SheetNumber, "£ #,##0.00");
                              sSheet.Set_Formula(RowNumber, i + 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:AK," + LookupValueBudget + ",FALSE),0)", SheetNumber, "£ #,##0.00");
                              sSheet.Set_Formula(RowNumber, i + 3, "=IFERROR((" + sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(i + 3) + (RowNumber + 1) + ")/" + sSheet.GetExcelColumnName(i + 3) + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");
                              sSheet.Set_Formula(RowNumber, i + 4, "=IFERROR((" + sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1) + ")/" + sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");

                              sSheet.Set_FontColour(sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1), Pink, null, SheetNumber);
                              sSheet.Set_FontColour(sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1), LightGreen, null, SheetNumber);
                              sSheet.Set_FontColour(sSheet.GetExcelColumnName(i + 3) + (RowNumber + 1), Color.Lavender, null, SheetNumber);

                              LookupValueBudget++;
                              LookupValueCurrentMonth += 3;
                           }

                           string Sum = "";
                           for (int i = 2; i < ((MonthColumnIndex - 1) * 5); i += 5)
                              Sum += sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1) + "+";
                           string Sum2024 = Sum.TrimEnd('+');

                           Sum = "";
                           for (int i = 1; i < ((MonthColumnIndex - 1) * 5); i += 5)
                              Sum += sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1) + "+";
                           string Sum2023 = Sum.TrimEnd('+');
                           Sum = "";
                           for (int i = 3; i < ((MonthColumnIndex - 1) * 5); i += 5)
                              Sum += sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1) + "+";
                           string Budget2024 = Sum.TrimEnd('+');

                           sSheet.Set_Formula(RowNumber, 61, "=" + Sum2024, SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 62, "=" + Sum2023, SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 63, "=" + Budget2024, SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 64, "=" + sSheet.GetExcelColumnName(62) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(63) + (RowNumber + 1), SheetNumber, "£ #,##0.00");
                           sSheet.Set_Formula(RowNumber, 65, "=" + sSheet.GetExcelColumnName(62) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(64) + (RowNumber + 1), SheetNumber, "£ #,##0.00");

                           RowNumber++;
                        }
                     }
                  }
               }

               for (int i = 4; i < 61; i += 5)
                  sSheet.Set_Colour_Gradient_Formatting(sSheet.GetExcelColumnName(i + 1) + "3:" + sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1), "45%", Color.Red, null, SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, "TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_FontColour("A" + (RowNumber + 1), Color.LightGray, null, SheetNumber);
               sSheet.Set_Bold_Range("A1:A" + (RowNumber + 1), true, SheetNumber);

               for (int i = 1; i < 61; i += 5)
               {
                  sSheet.Set_Formula(RowNumber, i, "=SUM(" + sSheet.GetExcelColumnName(i + 1) + "3:" + sSheet.GetExcelColumnName(i + 1) + (RowNumber) + ")", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, (i + 1), "=SUM(" + sSheet.GetExcelColumnName(i + 2) + "3:" + sSheet.GetExcelColumnName(i + 2) + (RowNumber) + ")", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, (i + 2), "=SUM(" + sSheet.GetExcelColumnName(i + 3) + "3:" + sSheet.GetExcelColumnName(i + 3) + (RowNumber) + ")", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, (i + 3), "=IFERROR((" + sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(i + 3) + (RowNumber + 1) + ")/" + sSheet.GetExcelColumnName(i + 3) + (RowNumber) + ",0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, (i + 4), "=IFERROR((" + sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1) + ")/" + sSheet.GetExcelColumnName(i + 1) + (RowNumber) + ",0)", SheetNumber, "£ #,##0");

                  sSheet.Set_FontColour(sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1), Pink, null, SheetNumber);
                  sSheet.Set_FontColour(sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1), LightGreen, null, SheetNumber);
                  sSheet.Set_FontColour(sSheet.GetExcelColumnName(i + 3) + (RowNumber + 1), Color.Lavender, null, SheetNumber);
                  sSheet.Set_FontColour(sSheet.GetExcelColumnName(i + 4) + (RowNumber + 1), Color.LightGray, null, SheetNumber);
                  sSheet.Set_FontColour(sSheet.GetExcelColumnName(i + 5) + (RowNumber + 1), Color.LightGray, null, SheetNumber);
               }

               sSheet.Set_Conditional_Formatting("BM3:BN" + (RowNumber + 1), ConditionalFormattingExpressionCondition.LessThan, "0", null, Color.Red, SheetNumber);

               sSheet.Set_FontColour("BJ2:BN" + (RowNumber + 1), Color.LightGray, null, SheetNumber);
               sSheet.Set_FontColour("BM3:BM" + (RowNumber + 1), Pink, null, SheetNumber);
               sSheet.Set_FontColour("BN3:BN" + (RowNumber + 1), Color.Lavender, null, SheetNumber);

               sSheet.Set_AllBorders("A1:BO" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_OutsideBorders("A1:BO" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("A" + (RowNumber + 1) + ":BO" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);

               sSheet.FreezePlanes(1, 0, SheetNumber);
               int StartIndex = ((MonthColumnIndex - 1) * 5) + 1;

               sSheet.Set_Column_Width(0, 59.00, "4.CUSTOMER £ v BDG v PY");

               for (int i = 1; i < 67; i++)
                  sSheet.Set_Column_Width(i, 13.57, "4.CUSTOMER £ v BDG v PY");

               if (StartIndex < 60)
                  sSheet.Hide_Columns(StartIndex, 60, SheetNumber);


               /**************************************************************************************************************************
               * CUSTOMERS NOT PURCHASED THIS YEAR THAT PURCHASED LAST YEAR
               *************************************************************************************************************************/
               RowNumber = 0;
               SheetNumber++;
               sSheet.Insert_Worksheet("CUSTOMERS NOT PURCHASED THIS YEAR THAT PURCHASED LAST YEAR", SheetNumber);

               sSheet.Set_Cell(RowNumber, 1, "CUSTOMERS NOT PURCHASED THIS YEAR THAT PURCHASED LAST YEAR", SheetNumber);

               sSheet.Merge_Cells("B1:C1", SheetNumber);
               sSheet.Set_Bold(RowNumber, 1, true, SheetNumber);
               sSheet.Set_Font_Size("B1", 12, SpreadsheetHorizontalAlignment.Center, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 1, "CUSTOMER", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "LAST INVOICE DATE", SheetNumber);

               sSheet.Set_Bold(RowNumber, 1, true, SheetNumber);
               sSheet.Set_Bold(RowNumber, 2, true, SheetNumber);
               sSheet.Set_Font_Size("B2", 11, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_Font_Size("C2", 11, SpreadsheetHorizontalAlignment.Center, SheetNumber);

               RowNumber++;

               DataTable PurchasedThisYear = new DataTable();
               DataTable PurchasedLastYear = new DataTable();
               for (int i = 3; i < ThisYearRowCount; i++)
               {
                  var Value = sSheet.Get_Cell_Text((i - 1), (((MonthColumnIndex - 1) * 3) - 2), Year + " MONTH SALES PER CUSTOMER");
                  if (Value == null || Value == "")
                  {
                     var Name = sSheet.Get_Cell_Text((i - 1), 0, Year + " MONTH SALES PER CUSTOMER");
                     if (Name != null && Name != "")
                     {
                        sqlstring = "SELECT tbl_Invoice.Invoice_Date " +
                                 "FROM tbl_Customer INNER JOIN " +
                                  "tbl_Invoice ON tbl_Customer.CustomerID = tbl_Invoice.CustomerID " +
                                 "WHERE(tbl_Invoice.Invoice_Date BETWEEN '" + LastYearStart + "' AND '" + LastYear + "-12-31" + "') AND(tbl_Customer.Name = N'" + Name + "') " +
                                 "ORDER BY tbl_Invoice.Invoice_Date ";

                        PurchasedLastYear = Invoices.RetrieveDataTable(sqlstring);

                        sqlstring = "SELECT tbl_Invoice.Invoice_Date " +
                                 "FROM tbl_Customer INNER JOIN " +
                                  "tbl_Invoice ON tbl_Customer.CustomerID = tbl_Invoice.CustomerID " +
                                 "WHERE(tbl_Invoice.Invoice_Date BETWEEN '" + Year + "-01-01" + "' AND '" + Year + "-12-31" + "') AND(tbl_Customer.Name = N'" + Name + "') " +
                                 "ORDER BY tbl_Invoice.Invoice_Date ";

                        PurchasedThisYear = Invoices.RetrieveDataTable(sqlstring);

                        if (PurchasedThisYear.Rows.Count == 0 && PurchasedLastYear.Rows.Count > 0)
                        {
                           int LastYearCount = (PurchasedLastYear.Rows.Count - 1);
                           string LastInvoiceDate = Classes.Global.ConvertToDateTime(PurchasedLastYear.Rows[LastYearCount].ItemArray[0]).ToString("dd-MM-yyyy");
                           sSheet.Set_Formula(RowNumber, 1, "='" + Year + " MONTH SALES PER CUSTOMER'!A" + i, SheetNumber);
                           sSheet.Set_Cell(RowNumber, 2, LastInvoiceDate, SheetNumber);

                           sSheet.Set_Font_Size("B" + (RowNumber + 1), 11, SpreadsheetHorizontalAlignment.Center, SheetNumber);
                           sSheet.Set_Font_Size("C" + (RowNumber + 1), 11, SpreadsheetHorizontalAlignment.Center, SheetNumber);

                           RowNumber++;
                        }
                        else
                        {
                        }
                     }
                  }
               }

               sSheet.Set_Column_Width(1, 46.43, SheetNumber);
               sSheet.Set_Column_Width(2, 46.43, SheetNumber);


               /**************************************************************************************************************************
               * SALES + WEIGHT EXPORT
               *************************************************************************************************************************/
               RowNumber = 0;
               SheetNumber++;
               sSheet.Add_Worksheet("SALES + WEIGHT EXPORT");

               sSheet.Set_Cell(RowNumber, 1, "VALUE", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "WEIGHT", SheetNumber);

               RowNumber++;

               foreach (DataRow row in YTDSales.Rows)
               {
                  if (Classes.Global.ConvertToString(row["Name"]) != "STANNAH STAIRLIFT EURO ACCOUNT")
                  {
                     if (Classes.Global.ConvertToString(row["Name"]) != "")
                     {
                        sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]), SheetNumber);
                        sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3) - 1) + ",FALSE),0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + ((MonthColumnIndex - 1) * 3) + ",FALSE),0)", SheetNumber, "#,##0");

                        RowNumber++;
                     }
                  }
               }

               sSheet.Auto_fit(0, 2, SheetNumber);

               /**************************************************************************************************************************
               * THIS YR VS LAST YR SALES VAR
               *************************************************************************************************************************/
               RowNumber = 0;
               SheetNumber++;
               sSheet.Insert_Worksheet(Year + " VS " + LastYear + " SALES VAR", SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, LastYear, SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 1, "ANNUAL TOTAL", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, Year, SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 5, "ANNUAL TOTAL", SheetNumber);
               sSheet.Set_Cell(RowNumber, 8, Year + " VAR TO " + LastYear, SheetNumber);
               sSheet.Set_Cell(RowNumber, 9, "TONNES", SheetNumber);
               sSheet.Set_Cell(RowNumber, 12, "REVENUE", SheetNumber);
               sSheet.Set_Cell(RowNumber, 15, "ASP/T", SheetNumber);

               sSheet.Set_Rotation("A1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("A1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_Bold_Range("A1:AN1", true, SheetNumber);
               sSheet.Merge_Cells("A1:A2", SheetNumber);
               sSheet.Set_Rotation("E1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("E1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Merge_Cells("E1:E2", SheetNumber);
               sSheet.Set_Rotation("I1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("I1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Merge_Cells("I1:I2", SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 1, "VALUE", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "WEIGHT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "ASP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "VALUE", SheetNumber);
               sSheet.Set_Cell(RowNumber, 6, "WEIGHT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 7, "ASP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 9, Year + " VS " + LastYear, SheetNumber);
               sSheet.Set_Cell(RowNumber, 10, "%VAR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 12, Year + " VS " + LastYear, SheetNumber);
               sSheet.Set_Cell(RowNumber, 13, "%VAR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 15, Year + " VS " + LastYear, SheetNumber);
               sSheet.Set_Cell(RowNumber, 16, "%VAR", SheetNumber);

               RowNumber++;

               sqlstring = "SELECT tbl_Customer.CustomerID, tbl_Customer.Account_Ref, LTRIM(RTRIM(tbl_Customer.Name)) AS Name, Inv.Line_Cost_Price, " +
                   "Inv.Line_Sale_Price, Inv.Line_Unit_Weight, Inv.Invoice_Month, tbl_Customer.Deleted " +
                   "FROM tbl_Customer LEFT OUTER JOIN(SELECT SUM(tbl_InvoiceItem.Cost_Price* tbl_InvoiceItem.Qty_Order) AS Line_Cost_Price, SUM(tbl_InvoiceItem.Net_Amount) AS Line_Sale_Price, " +
                   "SUM(tbl_Product.Unit_Weight * tbl_InvoiceItem.Qty_Order) AS Line_Unit_Weight, " +
                   "tbl_Invoice.CustomerID, MONTH(tbl_Invoice.Invoice_Date) AS Invoice_Month " +
                   "FROM tbl_Invoice AS tbl_Invoice LEFT OUTER JOIN " +
                   "tbl_Product RIGHT OUTER JOIN " +
                   "tbl_InvoiceItem ON tbl_Product.ProductID = tbl_InvoiceItem.ProductID ON tbl_Invoice.InvoiceID = tbl_InvoiceItem.InvoiceID " +
                   "WHERE(tbl_Invoice.Invoice_Date IS NULL OR " +
                   "tbl_Invoice.Invoice_Date BETWEEN'" + LastYear + "-01-01 00:00:00' AND '" + EndDate + "') " +
                   "GROUP BY tbl_Invoice.CustomerID, MONTH(Invoice_Date)) Inv ON tbl_Customer.CustomerID = Inv.CustomerID " +
           "WHERE (tbl_Customer.Deleted = 0 OR Inv.Line_Sale_Price > 0)" +
                   "ORDER BY tbl_Customer.Name ";


               List<string> Names = new List<string>();
               DataTable LastYearvsThisYearVarTable = Invoices.RetrieveDataTable(sqlstring, false);
               foreach (DataRow row in LastYearvsThisYearVarTable.Rows)
               {
                  Names.Add(row["Name"].ToString());
               }
               List<string> NewNames = Names.Distinct().ToList();

               LastYearvsThisYearVarTable.Dispose();
               LastYearvsThisYearVarTable = null;

               foreach (string Name in NewNames)
               {
                  if (Name != "STANNAH STAIRLIFT EURO ACCOUNT")
                  {
                     sSheet.Set_Cell(RowNumber, 8, Name, SheetNumber);
                     sSheet.Set_Formula(RowNumber, 0, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AZ,1,FALSE),\"\")", SheetNumber);
                     sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AZ,41,FALSE),0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AZ,42,FALSE),0)", SheetNumber, "#,##0");
                     sSheet.Set_Formula(RowNumber, 3, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AZ,43,FALSE),0)", SheetNumber, "£ #,##0.00");

                     sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,1,FALSE),\"\")", SheetNumber);
                     sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,38,FALSE),0)", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,39,FALSE),0)", SheetNumber, "#,##0");
                     sSheet.Set_Formula(RowNumber, 7, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,40,FALSE),0)", SheetNumber, "£ #,##0.00");


                     sSheet.Set_Formula(RowNumber, 9, "=G" + (RowNumber + 1) + "-C" + (RowNumber + 1), SheetNumber, "#,##0");
                     sSheet.Set_Formula(RowNumber, 10, "=IFERROR(J" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0.0 %");

                     sSheet.Set_Formula(RowNumber, 12, "=F" + (RowNumber + 1) + "-B" + (RowNumber + 1), SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 13, "=IFERROR(M" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0.0 %");

                     sSheet.Set_Formula(RowNumber, 15, "=H" + (RowNumber + 1) + "-D" + (RowNumber + 1), SheetNumber, "#,##0.00");
                     sSheet.Set_Formula(RowNumber, 16, "=IFERROR(P" + (RowNumber + 1) + "/D" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0.0 %");

                     RowNumber++;
                  }
               }
               sSheet.Set_Cell(RowNumber, 0, "TOTAL", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "=SUM(B3:B" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(C3:C" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(B" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ", 0)", SheetNumber, "£ #,##0.00");


               sSheet.Set_Cell(RowNumber, 4, "TOTAL", SheetNumber);
               sSheet.Set_Formula(RowNumber, 5, "=SUM(F3:F" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUM(G3:G" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(F" + (RowNumber + 1) + "/G" + (RowNumber + 1) + ", 0)", SheetNumber, "£ #,##0.00");

               sSheet.Set_Cell(RowNumber, 8, "TOTAL", SheetNumber);
               sSheet.Set_Formula(RowNumber, 9, "=SUM(J3:J" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 10, "=IFERROR(J" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0.0 %");

               sSheet.Set_Formula(RowNumber, 12, "=F" + (RowNumber + 1) + "-B" + (RowNumber + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 13, "=IFERROR(M" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0.0 %");

               sSheet.Set_Formula(RowNumber, 15, "=H" + (RowNumber + 1) + "-D" + (RowNumber + 1), SheetNumber, "#,##0.00");
               sSheet.Set_Formula(RowNumber, 16, "=IFERROR(P" + (RowNumber + 1) + "/D" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0.0 %");

               RowNumber++;

               sSheet.Set_Bold_Range("A" + RowNumber + ":Q" + RowNumber, true, SheetNumber);
               sSheet.Set_AllBorders("A1: K" + RowNumber, null, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("M1: N" + RowNumber, null, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("P1: Q" + RowNumber, null, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_FontColour("A1:A2", LightGreen, null, SheetNumber);
               sSheet.Set_FontColour("E1:E2", LightGreen, null, SheetNumber);
               sSheet.Set_FontColour("I1:J2", LightGreen, null, SheetNumber);
               sSheet.Set_FontColour("M1:M2", Color.Orange, null, SheetNumber);
               sSheet.Set_FontColour("P1:P2", Color.Plum, null, SheetNumber);
               sSheet.Set_FontColour("B1:D" + RowNumber, Color.LightGray, null, SheetNumber);
               sSheet.Set_FontColour("F1:H" + RowNumber, Color.LightGray, null, SheetNumber);

               sSheet.Set_Conditional_Formatting("J3:K" + RowNumber, ConditionalFormattingExpressionCondition.LessThan, "0", Color.Yellow, null, SheetNumber);
               sSheet.Set_Conditional_Formatting("M3:N" + RowNumber, ConditionalFormattingExpressionCondition.LessThan, "0", Color.Orange, null, SheetNumber);
               sSheet.Set_Conditional_Formatting("P3:Q" + RowNumber, ConditionalFormattingExpressionCondition.LessThan, "0", Color.Plum, null, SheetNumber);

               sSheet.Set_Column_Width(0, 46.43, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(1, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(2, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(3, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(4, 46.43, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(5, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(6, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(7, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(8, 46.43, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(9, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(10, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(12, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(13, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(15, 11.86, Year + " VS " + LastYear + " SALES VAR");
               sSheet.Set_Column_Width(16, 11.86, Year + " VS " + LastYear + " SALES VAR");

               /**************************************************************************************************************************
               * THIS YR BUDGET VS THIS YR ACTUAL VAR
               *************************************************************************************************************************/
               RowNumber = 0;
               SheetNumber++;

               sSheet.Insert_Worksheet(Year + " BUDGET VS " + Year + " ACTUALS", SheetNumber);
               CustomervBdgvPy = SheetNumber;

               sSheet.Set_Cell(RowNumber, 0, Year + " BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 1, "ANNUAL TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 4, Year + " ACTUAL", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 5, "ANNUAL TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 8, "BUDGET VAR TO ACTUAL", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 9, "REVENUE", SheetNumber, SpreadsheetHorizontalAlignment.Center);

               sSheet.Set_Rotation("A1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("A1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_Bold_Range("A1:AN1", true, SheetNumber);
               sSheet.Merge_Cells("A1:A2", SheetNumber);
               sSheet.Set_Rotation("E1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("E1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Merge_Cells("E1:E2", SheetNumber);
               sSheet.Set_Rotation("I1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("I1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Merge_Cells("I1:I2", SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 1, "VALUE", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, "WEIGHT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "ASP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "VALUE", SheetNumber);
               sSheet.Set_Cell(RowNumber, 6, "WEIGHT", SheetNumber);
               sSheet.Set_Cell(RowNumber, 7, "ASP", SheetNumber);
               sSheet.Set_Cell(RowNumber, 9, Year + " VAR", SheetNumber);
               sSheet.Set_Cell(RowNumber, 10, "%VAR", SheetNumber);

               RowNumber++;

               for (int i = 3; i < ThisYearRowCount; i++)
               {
                  sSheet.Set_Formula(RowNumber, 0, "='" + Year + " MONTH SALES PER CUSTOMER'!A" + i, SheetNumber);
                  sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:AZ,14,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 4, "='" + Year + " MONTH SALES PER CUSTOMER'!A" + i, SheetNumber);
                  sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,38,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,39,FALSE),0)", SheetNumber, "#,##0");
                  sSheet.Set_Formula(RowNumber, 7, "=IFERROR(VLOOKUP(I" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,40,FALSE),0)", SheetNumber, "£ #,##0.00");
                  sSheet.Set_Formula(RowNumber, 8, "='" + Year + " MONTH SALES PER CUSTOMER'!A" + i, SheetNumber);
                  sSheet.Set_Formula(RowNumber, 9, "=IFERROR(F" + (RowNumber + 1) + "-B" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 10, "=IFERROR(J" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0.0 %");

                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 0, "TOTAL", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "=SUM(B3:B" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUM(C3:C" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(B" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ", 0)", SheetNumber, "£ #,##0.00");


               sSheet.Set_Cell(RowNumber, 4, "TOTAL", SheetNumber);
               sSheet.Set_Formula(RowNumber, 5, "=SUM(F3:F" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUM(G3:G" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(F" + (RowNumber + 1) + "/G" + (RowNumber + 1) + ", 0)", SheetNumber, "£ #,##0.00");

               sSheet.Set_Cell(RowNumber, 8, "TOTAL", SheetNumber);
               sSheet.Set_Formula(RowNumber, 9, "=SUM(J3:J" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 10, "=IFERROR(J" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0.0 %");

               RowNumber++;


               sSheet.Set_AllBorders("A1:K" + RowNumber, null, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_Conditional_Formatting("J3:K" + RowNumber, ConditionalFormattingExpressionCondition.LessThan, "0", null, Color.Red, SheetNumber);
               sSheet.Set_FontColour("A1:J2", LightGreen, null, SheetNumber);
               sSheet.Set_FontColour("B1:D" + RowNumber, Color.LightGray, null, SheetNumber);
               sSheet.Set_FontColour("F1:H" + RowNumber, Color.LightGray, null, SheetNumber);

               sSheet.Set_Column_Width(0, 46.43, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(1, 11.86, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(2, 11.86, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(3, 11.86, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(4, 46.43, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(5, 11.86, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(6, 11.86, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(7, 11.86, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(8, 46.43, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(9, 11.86, Year + " BUDGET VS " + Year + " ACTUALS");
               sSheet.Set_Column_Width(10, 11.86, Year + " BUDGET VS " + Year + " ACTUALS");

               /**************************************************************************************************************************
               * MONTH SALES £ & KG
               *************************************************************************************************************************/

               SheetNumber++;
               RowNumber = 0;

               sSheet.Insert_Worksheet("5." + Month + " SALES £&KG", SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, Month, SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 1, Month + " BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 5, Month + " ACTUAL", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 9, Month + " ACTUAL", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 13, "YTD BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 17, "YTD ACTUAL", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 21, "YTD ACTUAL", SheetNumber, SpreadsheetHorizontalAlignment.Center);

               sSheet.Set_Rotation("A1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("A1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Set_Bold_Range("A1:AN1", true, SheetNumber);
               sSheet.Merge_Cells("A1:A2", SheetNumber);
               sSheet.Merge_Cells("B1:D1", SheetNumber);
               sSheet.Merge_Cells("F1:H1", SheetNumber);
               sSheet.Merge_Cells("J1:L1", SheetNumber);
               sSheet.Merge_Cells("N1:P1", SheetNumber);
               sSheet.Merge_Cells("R1:T1", SheetNumber);
               sSheet.Merge_Cells("V1:X1", SheetNumber);
               sSheet.Set_BackColour("A1:A2", LightGreen, SheetNumber);
               sSheet.Set_BackColour("B1:D1", Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("F1:H1", Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("J1:L1", Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("N1:P1", Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("R1:T1", Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("V1:X1", Color.LightGray, SheetNumber);

               RowNumber++;

               for (int i = 1; i <= 21; i += 4)
               {
                  sSheet.Set_Cell(RowNumber, i, "VALUE", SheetNumber);
                  sSheet.Set_Cell(RowNumber, i + 1, "WEIGHT", SheetNumber);
                  sSheet.Set_Cell(RowNumber, i + 2, "ASP", SheetNumber);
               }

               sSheet.Set_BackColour("B2:D2", LightGreen, SheetNumber);
               sSheet.Set_BackColour("F2:H2", LightGreen, SheetNumber);
               sSheet.Set_BackColour("J2:L2", LightGreen, SheetNumber);
               sSheet.Set_BackColour("N2:P2", LightGreen, SheetNumber);
               sSheet.Set_BackColour("R2:T2", LightGreen, SheetNumber);
               sSheet.Set_BackColour("V2:X2", LightGreen, SheetNumber);

               RowNumber++;

               foreach (DataRow row in YTDSales.Rows)
               {
                  bool Outdated = OutdatedList.Contains(Classes.Global.ConvertToString(row["Name"]).Trim());
                  bool inBudget = BudgetNameList.Contains(Classes.Global.ConvertToString(row["Name"]).Trim());
                  if (!Outdated || inBudget)
                  {
                     if (Classes.Global.ConvertToString(row["Name"]) != "STANNAH STAIRLIFT EURO ACCOUNT")
                     {
                        if (Classes.Global.ConvertToString(row["Name"]) != "")
                        {
                           sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:AZ," + (MonthColumnIndex) + ",FALSE),0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", 'FORECAST WITH WEIGHT'!A:AZ," + (((MonthColumnIndex) + (MonthColumnIndex)) - 1) + ",FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Formula(RowNumber, 3, "=IFERROR(B" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

                           sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3) - 1) + ",FALSE),0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", '" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + ((MonthColumnIndex - 1) * 3) + ",FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Formula(RowNumber, 7, "=IFERROR(F" + (RowNumber + 1) + "/G" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

                           sSheet.Set_Formula(RowNumber, 9, "=IFERROR(F" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
                           sSheet.Set_Formula(RowNumber, 10, "=IFERROR(G" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
                           sSheet.Set_Formula(RowNumber, 11, "=IFERROR(H" + (RowNumber + 1) + "/D" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           sSheet.Set_Formula(RowNumber, 13, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:AZ,14,FALSE),0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 14, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", 'FORECAST WITH WEIGHT'!A:AZ,27,FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Formula(RowNumber, 15, "=IFERROR(N" + (RowNumber + 1) + "/O" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

                           sSheet.Set_Formula(RowNumber, 17, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,41,FALSE),0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 18, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", '" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,42,FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Formula(RowNumber, 19, "=IFERROR(R" + (RowNumber + 1) + "/S" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

                           sSheet.Set_Formula(RowNumber, 21, "=IFERROR(R" + (RowNumber + 1) + "/N" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
                           sSheet.Set_Formula(RowNumber, 22, "=IFERROR(S" + (RowNumber + 1) + "/O" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
                           sSheet.Set_Formula(RowNumber, 23, "=IFERROR(T" + (RowNumber + 1) + "/P" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           RowNumber++;
                        }
                     }
                  }
               }

               sSheet.Set_Cell(RowNumber, 0, "New / OTHER", SheetNumber);


               sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", 'FORECAST WITH WEIGHT'!A:AZ," + (((MonthColumnIndex) + (MonthColumnIndex)) - 2) + ",FALSE),0)", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", 'FORECAST WITH WEIGHT'!A:AZ," + (((MonthColumnIndex) + (MonthColumnIndex)) - 1) + ",FALSE),0)", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(B" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "TOTAL:", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=SUBTOTAL(109,B3:B" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUBTOTAL(109,C3:C" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(B" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

               sSheet.Set_Formula(RowNumber, 5, "=SUBTOTAL(109,F3:F" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUBTOTAL(109,G3:G" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(F" + (RowNumber + 1) + "/G" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

               sSheet.Set_Formula(RowNumber, 9, "=IFERROR(F" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
               sSheet.Set_Formula(RowNumber, 10, "=IFERROR(G" + (RowNumber + 1) + "/C" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
               sSheet.Set_Formula(RowNumber, 11, "=IFERROR(H" + (RowNumber + 1) + "/D" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

               sSheet.Set_Formula(RowNumber, 13, "=SUBTOTAL(109,N3:N" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 14, "=SUBTOTAL(109,O3:O" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 15, "=IFERROR(N" + (RowNumber + 1) + "/O" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

               sSheet.Set_Formula(RowNumber, 17, "=SUBTOTAL(109,R3:R" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 18, "=SUBTOTAL(109,S3:S" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 19, "=IFERROR(R" + (RowNumber + 1) + "/S" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

               sSheet.Set_Formula(RowNumber, 21, "=IFERROR(R" + (RowNumber + 1) + "/N" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
               sSheet.Set_Formula(RowNumber, 22, "=IFERROR(S" + (RowNumber + 1) + "/O" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
               sSheet.Set_Formula(RowNumber, 23, "=IFERROR(T" + (RowNumber + 1) + "/P" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

               sSheet.Set_BackColour("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("F" + (RowNumber + 1) + ":H" + (RowNumber + 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("J" + (RowNumber + 1) + ":L" + (RowNumber + 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("N" + (RowNumber + 1) + ":P" + (RowNumber + 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("R" + (RowNumber + 1) + ":T" + (RowNumber + 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("V" + (RowNumber + 1) + ":X" + (RowNumber + 1), Color.LightGray, SheetNumber);

               sSheet.Set_AllBorders("A1:A" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("B1:D" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("F1:H" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("J1:L" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("N1:P" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("R1:T" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("V1:X" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);

               sSheet.Set_OutsideBorders("A1:A" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("B1:D" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("F1:H" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("J1:L" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("N1:P" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("R1:T" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("V1:X" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);

               sSheet.Set_Column_Width(0, 46.43, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(1, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(2, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(3, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(5, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(6, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(7, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(9, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(10, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(11, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(13, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(14, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(15, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(17, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(18, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(19, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(21, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(22, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Set_Column_Width(23, 11.86, "5." + Month + " SALES £&KG");
               sSheet.Auto_Filter("A2:X2", SheetNumber);


               /**************************************************************************************************************************
               * SG INTERNAL SALES
               *************************************************************************************************************************/
               Color LightBlue = ColorTranslator.FromHtml("#33CCCC");

               SheetNumber++;
               RowNumber = 0;

               sSheet.Insert_Worksheet("6.SG INTERNAL SALES", SheetNumber);

               sSheet.Set_Cell(RowNumber, 1, "REVENUE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 5, "VOLUME", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 9, "STOCK", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 13, "PURCHASES", SheetNumber, SpreadsheetHorizontalAlignment.Center);

               sSheet.Merge_Cells("B1:D1", SheetNumber);
               sSheet.Merge_Cells("F1:H1", SheetNumber);
               sSheet.Merge_Cells("J1:L1", SheetNumber);
               sSheet.Merge_Cells("N1:O1", SheetNumber);
               sSheet.Set_BackColour("B1:D2", Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("F1:H2", Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("J1:L2", Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("N1:O2", Color.LightGray, SheetNumber);
               sSheet.Merge_Cells("A1:A2", SheetNumber);
               sSheet.Set_Bold_Range("A1:P2", true, SheetNumber);
               sSheet.Set_BackColour("A1:A2", LightBlue, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 1, "BUDGET", SheetNumber);
               sSheet.Set_Cell(RowNumber, 2, Month + " SALES", SheetNumber);
               sSheet.Set_Cell(RowNumber, 3, "% OF BUDGET", SheetNumber);
               sSheet.Set_Cell(RowNumber, 5, "BUDGET", SheetNumber);
               sSheet.Set_Cell(RowNumber, 6, Month + " SALES", SheetNumber);
               sSheet.Set_Cell(RowNumber, 7, "% OF BUDGET", SheetNumber);
               sSheet.Set_Cell(RowNumber, 9, "TARGET STOCK", SheetNumber);
               sSheet.Set_Cell(RowNumber, 10, "STOCK KGS", SheetNumber);
               sSheet.Set_Cell(RowNumber, 11, "% ", SheetNumber);
               sSheet.Set_Cell(RowNumber, 13, "PURCHASES KGS", SheetNumber);
               sSheet.Set_Cell(RowNumber, 14, "% OF SALES", SheetNumber);

               sSheet.Set_OutsideBorders("B" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("F" + (RowNumber + 1) + ":H" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("J" + (RowNumber + 1) + ":L" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("N" + (RowNumber + 1) + ":O" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);

               RowNumber++;

               sqlstring = "SELECT tbl_Customer.Name, SUM(View_Stock_Theoretical_Weight.Unit_Weight * View_Stock_Theoretical_Weight.Qty_Free) AS Weight, tbl_Customer.CustomerID " +
                        "FROM tbl_Customer LEFT OUTER JOIN " +
                        "View_Stock_Theoretical_Weight ON tbl_Customer.CustomerID = View_Stock_Theoretical_Weight.Default_Customer " +
                        "WHERE(tbl_Customer.Deleted = 0) " +
                        "GROUP BY tbl_Customer.Name, tbl_Customer.CustomerID " +
                        "ORDER BY tbl_Customer.Name";

               DataTable theoStockdt = Invoices.RetrieveDataTable(sqlstring);

               List<TheoStockModel> TheoStockList = theoStockdt.AsEnumerable().Select(x => new TheoStockModel
               {
                  Name = x.Field<string>("Name"),
                  StockKGs = x.Field<double?>("Weight"),
                  CustomerID = x.Field<string>("CustomerID")

               }).ToList();

               theoStockdt.Dispose();
               theoStockdt = null;

               foreach (TheoStockModel stockModel in TheoStockList)
               {
                  if (stockModel.StockKGs == null)
                     stockModel.StockKGs = 0;

                  bool Outdated = OutdatedList.Contains(Classes.Global.ConvertToString(stockModel.Name).Trim());
                  bool inBudget = BudgetNameList.Contains(Classes.Global.ConvertToString(stockModel.Name).Trim());
                  if (!Outdated || inBudget)
                  {
                     bool newNoBudget = NewnoBudgetList.Contains(stockModel.Name.Trim());
                     if (!newNoBudget)
                     {
                        if (stockModel.Name == "STANNAH STAIRLIFT EURO ACCOUNT")
                        {
                        }
                        else if (stockModel.Name == "STANNAH STAIRLIFTS LTD")
                        {
                           TheoStockModel StanEuroModel = TheoStockList.Where(x => x.Name == "STANNAH STAIRLIFT EURO ACCOUNT").FirstOrDefault();

                           sSheet.Set_Cell(RowNumber, 0, stockModel.Name.Trim(), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:AZ," + (MonthColumnIndex) + ",FALSE),0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3) - 1) + ",FALSE),0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 3, "=IFERROR(C" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", 'FORECAST WITH WEIGHT'!A:AZ," + (((MonthColumnIndex) + (MonthColumnIndex)) - 1) + ",FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", '" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + ((MonthColumnIndex - 1) * 3) + ",FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Formula(RowNumber, 7, "=IFERROR(G" + (RowNumber + 1) + "/F" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", 'FORECAST ASP CALC'!A:AZ,5,FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Cell(RowNumber, 10, Math.Round(Classes.Global.ConvertToDouble(stockModel.StockKGs + StanEuroModel.StockKGs), 0, MidpointRounding.ToEven), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 11, "=IFERROR(K" + (RowNumber + 1) + "/J" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           sqlstring = "SELECT SUM(tbl_StockPack.Qty_Original * tbl_Product.Unit_Weight_Without_Components) AS Purchases " +
                              "FROM tbl_StockPack INNER JOIN " +
                              "tbl_Product ON tbl_StockPack.ProductID = tbl_Product.ProductID INNER JOIN " +
                              "tbl_PurchaseOrderItem ON tbl_StockPack.PurchaseOrderItemID = tbl_PurchaseOrderItem.PurchaseOrderItemID INNER JOIN " +
                              "tbl_PurchaseOrder ON tbl_PurchaseOrderItem.PurchaseOrderID = tbl_PurchaseOrder.PurchaseOrderID INNER JOIN " +
                              "tbl_Supplier ON tbl_PurchaseOrder.SupplierID = tbl_Supplier.SupplierID " +
                              "WHERE (tbl_StockPack.OriginalPackID IS NULL) AND (tbl_Product.Default_Customer = N'" + stockModel.CustomerID + "') AND " +
                              "(tbl_Supplier.Account_Ref = N'CORTIZO' OR tbl_Supplier.Account_Ref = N'SMARTALU' OR tbl_Supplier.Account_Ref = N'ORIGIN') AND " +
                              "(tbl_StockPack.Pack_Date BETWEEN '" + StartDate + "' AND '" + EndDate + "')";

                           DataTable pdt = Invoices.RetrieveDataTable(sqlstring);
                           PurchasesKGsModel purchase = pdt.AsEnumerable().Select(s => new PurchasesKGsModel { PurchasesKgs = s.Field<double?>("Purchases") }).FirstOrDefault();

                           if (purchase.PurchasesKgs == null)
                              purchase.PurchasesKgs = 0;

                           sqlstring = "SELECT SUM(tbl_StockPack.Qty_Original * tbl_Product.Unit_Weight_Without_Components) AS Purchases " +
                              "FROM tbl_StockPack INNER JOIN " +
                              "tbl_Product ON tbl_StockPack.ProductID = tbl_Product.ProductID INNER JOIN " +
                              "tbl_PurchaseOrderItem ON tbl_StockPack.PurchaseOrderItemID = tbl_PurchaseOrderItem.PurchaseOrderItemID INNER JOIN " +
                              "tbl_PurchaseOrder ON tbl_PurchaseOrderItem.PurchaseOrderID = tbl_PurchaseOrder.PurchaseOrderID INNER JOIN " +
                              "tbl_Supplier ON tbl_PurchaseOrder.SupplierID = tbl_Supplier.SupplierID " +
                              "WHERE (tbl_StockPack.OriginalPackID IS NULL) AND (tbl_Product.Default_Customer = N'" + StanEuroModel.CustomerID + "') AND " +
                              "(tbl_Supplier.Account_Ref = N'CORTIZO' OR tbl_Supplier.Account_Ref = N'SMARTALU' OR tbl_Supplier.Account_Ref = N'ORIGIN') AND " +
                              "(tbl_StockPack.Pack_Date BETWEEN '" + StartDate + "' AND '" + EndDate + "')";

                           pdt = Invoices.RetrieveDataTable(sqlstring);
                           PurchasesKGsModel StanEuropurchase = pdt.AsEnumerable().Select(s => new PurchasesKGsModel { PurchasesKgs = s.Field<double?>("Purchases") }).FirstOrDefault();

                           if (StanEuropurchase.PurchasesKgs == null)
                              StanEuropurchase.PurchasesKgs = 0;

                           sSheet.Set_Cell(RowNumber, 13, Math.Round(Classes.Global.ConvertToDouble(purchase.PurchasesKgs + StanEuropurchase.PurchasesKgs), 0, MidpointRounding.ToEven), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 14, "=IFERROR(N" + (RowNumber + 1) + "/G" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           RowNumber++;
                        }
                        else
                        {
                           sSheet.Set_Cell(RowNumber, 0, stockModel.Name.Trim(), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:AZ," + (MonthColumnIndex) + ",FALSE),0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3) - 1) + ",FALSE),0)", SheetNumber, "£ #,##0");
                           sSheet.Set_Formula(RowNumber, 3, "=IFERROR(C" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", 'FORECAST WITH WEIGHT'!A:AZ," + (((MonthColumnIndex) + (MonthColumnIndex)) - 1) + ",FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", '" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + ((MonthColumnIndex - 1) * 3) + ",FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Formula(RowNumber, 7, "=IFERROR(G" + (RowNumber + 1) + "/F" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ", 'FORECAST ASP CALC'!A:AZ,5,FALSE),0)", SheetNumber, "#,##0");
                           sSheet.Set_Cell(RowNumber, 10, Math.Round(Classes.Global.ConvertToDouble(stockModel.StockKGs), 0, MidpointRounding.ToEven), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 11, "=IFERROR(K" + (RowNumber + 1) + "/J" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           sqlstring = "SELECT SUM(tbl_StockPack.Qty_Original * tbl_Product.Unit_Weight_Without_Components) AS Purchases " +
                              "FROM tbl_StockPack INNER JOIN " +
                              "tbl_Product ON tbl_StockPack.ProductID = tbl_Product.ProductID INNER JOIN " +
                              "tbl_PurchaseOrderItem ON tbl_StockPack.PurchaseOrderItemID = tbl_PurchaseOrderItem.PurchaseOrderItemID INNER JOIN " +
                              "tbl_PurchaseOrder ON tbl_PurchaseOrderItem.PurchaseOrderID = tbl_PurchaseOrder.PurchaseOrderID INNER JOIN " +
                              "tbl_Supplier ON tbl_PurchaseOrder.SupplierID = tbl_Supplier.SupplierID " +
                              "WHERE (tbl_StockPack.OriginalPackID IS NULL) AND (tbl_Product.Default_Customer = N'" + stockModel.CustomerID + "') AND " +
                              "(tbl_Supplier.Account_Ref = N'CORTIZO' OR tbl_Supplier.Account_Ref = N'SMARTALU' OR tbl_Supplier.Account_Ref = N'ORIGIN') AND " +
                              "(tbl_StockPack.Pack_Date BETWEEN '" + StartDate + "' AND '" + EndDate + "')";

                           DataTable pdt = Invoices.RetrieveDataTable(sqlstring);
                           PurchasesKGsModel purchase = pdt.AsEnumerable().Select(s => new PurchasesKGsModel { PurchasesKgs = s.Field<double?>("Purchases") }).FirstOrDefault();

                           if (purchase.PurchasesKgs == null)
                              purchase.PurchasesKgs = 0;

                           sSheet.Set_Cell(RowNumber, 13, Math.Round(Classes.Global.ConvertToDouble(purchase.PurchasesKgs), 0, MidpointRounding.ToEven), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 14, "=IFERROR(N" + (RowNumber + 1) + "/G" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

                           RowNumber++;
                        }
                     }
                  }
               }
               sSheet.Set_BackColour("A3:A" + RowNumber, Color.LightGray, SheetNumber);

               sSheet.Set_Cell(RowNumber, 1, "", SheetNumber);
               sSheet.Set_Formula(RowNumber, 2, "=SUBTOTAL(109,C3:C" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Cell(RowNumber, 3, "", SheetNumber);

               sSheet.Set_Formula(RowNumber, 5, "=SUBTOTAL(109,F3:F" + (RowNumber) + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUBTOTAL(109,G3:G" + (RowNumber) + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(G" + (RowNumber + 1) + "/F" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

               sSheet.Set_Formula(RowNumber, 9, "=SUBTOTAL(109,J3:J" + (RowNumber) + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 10, "=SUBTOTAL(109,K3:K" + (RowNumber) + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 11, "=IFERROR(K" + (RowNumber + 1) + "/J" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

               sSheet.Set_Formula(RowNumber, 13, "=SUBTOTAL(109,N3:N" + (RowNumber) + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 14, "=IFERROR(N" + (RowNumber + 1) + "/G" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

               sSheet.Set_Bold_Range("B" + (RowNumber + 1) + ":O" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_AllBorders("A1:D" + (RowNumber), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("F1:H" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("J1:L" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("N1:O" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_OutsideBorders("A1:A" + (RowNumber), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("B1:D" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("F1:H" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("J1:L" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("N1:O" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);

               sSheet.Set_OutsideBorders("B" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("F" + (RowNumber + 1) + ":H" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("J" + (RowNumber + 1) + ":L" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("N" + (RowNumber + 1) + ":O" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);

               RowNumber += 2;

               sSheet.Set_BackColour("A" + (RowNumber - 1) + ":A" + (RowNumber + 1), LightBlue, SheetNumber);
               sSheet.Set_BackColour("B" + (RowNumber - 1) + ":D" + (RowNumber - 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("F" + (RowNumber - 1) + ":H" + (RowNumber - 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("J" + (RowNumber - 1) + ":L" + (RowNumber - 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("N" + (RowNumber - 1) + ":O" + (RowNumber - 1), Color.LightGray, SheetNumber);

               sSheet.Set_BackColour("B" + (RowNumber) + ":O" + (RowNumber), LightBlue, SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, "New / OTHER", SheetNumber);
               sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'FORECAST WITH WEIGHT'!A:AZ," + (((MonthColumnIndex) + (MonthColumnIndex)) - 2) + ",FALSE),0)", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 2, "=SUBTOTAL(109,C" + (RowNumber + 2) + ":C" + (RowNumber + 1 + NewnoBudgetList.Count) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(C" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
               sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'FORECAST WITH WEIGHT'!A:AZ," + (((MonthColumnIndex) + (MonthColumnIndex)) - 1) + ",FALSE),0)", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 6, "=SUBTOTAL(109,G" + (RowNumber + 2) + ":G" + (RowNumber + 1 + NewnoBudgetList.Count) + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(G" + (RowNumber + 1) + "/F" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");
               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":H" + (RowNumber + 1), true, SheetNumber);
               int NewOtherRow = RowNumber;
               RowNumber++;
               foreach (string Name in NewnoBudgetList.OrderBy(o => o))
               {
                  sSheet.Set_Cell(RowNumber, 0, Name, SheetNumber);
                  sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3) - 1) + ",FALSE),0)", SheetNumber, "£ #,##0");

                  sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'FORECAST WITH WEIGHT'!A:AZ," + (((MonthColumnIndex) + (MonthColumnIndex)) - 1) + ",FALSE),0)", SheetNumber, "#,##0");
                  sSheet.Set_Formula(RowNumber, 6, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3)) + ",FALSE),0)", SheetNumber, "#,##0");

                  RowNumber++;
               }
               sSheet.Set_BackColour("A" + (NewOtherRow + 1) + ":A" + (RowNumber + 1), Color.LightGray, SheetNumber);
               sSheet.Set_AllBorders("A" + (NewOtherRow + 1) + ":D" + (RowNumber), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("F" + (NewOtherRow + 1) + ":H" + (RowNumber), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_OutsideBorders("A" + (NewOtherRow + 1) + ":A" + (RowNumber), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("B" + (NewOtherRow + 1) + ":D" + (RowNumber), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("F" + (NewOtherRow + 1) + ":H" + (RowNumber), Color.Black, SheetNumber, BorderLineStyle.Medium);

               RowNumber++;

               sSheet.Set_BackColour("A" + RowNumber, LightBlue, SheetNumber);
               sSheet.Set_BackColour("B" + (RowNumber) + ":O" + (RowNumber), LightBlue, SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, "TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=SUBTOTAL(109,B3:B" + (RowNumber) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 2, "=C" + (NewOtherRow - 1) + "+C" + (NewOtherRow + 1), SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR(C" + (RowNumber + 1) + "/B" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

               sSheet.Set_Formula(RowNumber, 5, "=F" + (NewOtherRow - 1) + "+F" + (NewOtherRow + 1), SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 6, "=G" + (NewOtherRow - 1) + "+G" + (NewOtherRow + 1), SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 7, "=IFERROR(G" + (RowNumber + 1) + "/F" + (RowNumber + 1) + ",0)", SheetNumber, "% #,##0");

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":H" + (RowNumber + 1), true, SheetNumber);
               sSheet.Set_BackColour("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.LightGray, SheetNumber);
               sSheet.Set_BackColour("F" + (RowNumber + 1) + ":H" + (RowNumber + 1), Color.LightGray, SheetNumber);

               sSheet.Set_AllBorders("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_AllBorders("F" + (RowNumber + 1) + ":H" + (RowNumber + 1), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_OutsideBorders("A" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("B" + (RowNumber + 1) + ":D" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("F" + (RowNumber + 1) + ":H" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);

               sSheet.Set_BackColour("J" + (NewOtherRow + 1) + ":O" + (RowNumber + 1), LightBlue, SheetNumber);


               sSheet.Set_BackColour("E1:E" + (RowNumber + 1), LightBlue, SheetNumber);
               sSheet.Set_BackColour("I1:I" + (RowNumber + 1), LightBlue, SheetNumber);
               sSheet.Set_BackColour("M1:M" + (RowNumber + 1), LightBlue, SheetNumber);
               sSheet.Set_BackColour("P1:P" + (RowNumber + 1), LightBlue, SheetNumber);

               sSheet.Set_Font_Size("A1:O" + (RowNumber + 1), 12, SheetNumber);
               sSheet.Auto_fit(1, 16, SheetNumber);
               sSheet.Auto_Filter("A2:O2", SheetNumber);

               sSheet.Set_Conditional_Formatting("D3: D" + (NewOtherRow - 2), ConditionalFormattingExpressionCondition.GreaterThan, "0.99", Color.PaleGreen, Color.Green, SheetNumber);
               sSheet.Set_Conditional_Formatting("H3: H" + (NewOtherRow - 2), ConditionalFormattingExpressionCondition.GreaterThan, "0.99", Color.PaleGreen, Color.Green, SheetNumber);
               sSheet.Set_Conditional_Formatting("L3: L" + (NewOtherRow - 2), ConditionalFormattingExpressionCondition.GreaterThan, "1.49", Color.MistyRose, Color.Red, SheetNumber);
               sSheet.Set_Conditional_Formatting("O3: O" + (NewOtherRow - 2), ConditionalFormattingExpressionCondition.GreaterThan, "1.49", Color.MistyRose, Color.Red, SheetNumber);

               /**************************************************************************************************************************
               * YTD LOWER THAN BUDGET
               *************************************************************************************************************************/

               SheetNumber++;
               RowNumber = 0;

               sSheet.Insert_Worksheet("YTD BELOW BUDGET", SheetNumber);

               sSheet.Set_Cell(RowNumber, 1, "TOTALS", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Merge_Cells("B1:D1", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "TONNAGE TOTALS", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Merge_Cells("E1:G1", SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "THIS YR V BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Rotation("A2", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("A2", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Merge_Cells("A1:A2", SheetNumber);

               sSheet.Set_Cell(RowNumber, 1, Year + " YTD", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 2, "BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 3, "VAR AGAINST BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 4, Year + " YTD TONNAGE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 5, "BUDGET TONNAGE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 6, "VAR AGAINST BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Center);

               sSheet.Set_FontColour("A1:A2", LightGreen, Color.Black, SheetNumber);
               sSheet.Set_FontColour("B1:G2", Color.LightGray, Color.Black, SheetNumber);

               sSheet.Set_Bold_Range("A1:G2", true, SheetNumber);

               RowNumber++;

               foreach (DataRow row in YTDSales.Rows)
               {
                  bool Outdated = OutdatedList.Contains(Classes.Global.ConvertToString(row["Name"]).Trim());
                  bool inBudget = BudgetNameList.Contains(Classes.Global.ConvertToString(row["Name"]).Trim());
                  if (!Outdated || inBudget)
                  {
                     if (Classes.Global.ConvertToString(row["Name"]) != "")
                     {
                        if (Classes.Global.ConvertToString(row["Name"]) != "STANNAH STAIRLIFT EURO ACCOUNT")
                        {
                           sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]).Trim(), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:BZ,38,FALSE),0)", SheetNumber, "£ #,##0.00");
                           sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:AK,14,FALSE),0)", SheetNumber, "£ #,##0.00");
                           sSheet.Set_Formula(RowNumber, 3, "=IFERROR((B" + (RowNumber + 1) + "- C" + (RowNumber + 1) + ")/C" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");
                           sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:BZ,39,FALSE),0)", SheetNumber, "#,##0.00");
                           sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'FORECAST WITH WEIGHT'!A:BZ,27,FALSE),0)", SheetNumber, "#,##0.00");
                           sSheet.Set_Formula(RowNumber, 6, "=IFERROR((E" + (RowNumber + 1) + "- F" + (RowNumber + 1) + ")/F" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");

                           double YTDValue = Classes.Global.ConvertToDouble(sSheet.Get_Cell_Text_From_Formula(sSheet.Get_Row_Index_From_Name(Classes.Global.ConvertToString(row["Name"]), Year + " MONTH SALES PER CUSTOMER"), 37, Year + " MONTH SALES PER CUSTOMER").ToString().Replace("£", ""));
                           double BudgetValue = Classes.Global.ConvertToDouble(sSheet.Get_Cell_Text_From_Formula(sSheet.Get_Row_Index_From_Name(Classes.Global.ConvertToString(row["Name"]), Year + " BUDGET"), 13, Year + " BUDGET").ToString().Replace("£", ""));

                           if (YTDValue < BudgetValue)
                              RowNumber++;
                        }
                     }
                  }
               }

               sSheet.Set_Cell(RowNumber, 0, "TOTAL:", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=SUBTOTAL(9, B3:B" + RowNumber + ")", SheetNumber, "£ #,##0.00");
               sSheet.Set_Formula(RowNumber, 2, "=SUBTOTAL(9, C3:C" + RowNumber + ")", SheetNumber, "£ #,##0.00");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR((B" + (RowNumber + 1) + "- C" + (RowNumber + 1) + ")/C" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=SUBTOTAL(9, E3:E" + RowNumber + ")", SheetNumber, "£ #,##0.00");
               sSheet.Set_Formula(RowNumber, 5, "=SUBTOTAL(9, F3:F" + RowNumber + ")", SheetNumber, "£ #,##0.00");
               sSheet.Set_Formula(RowNumber, 6, "=IFERROR((E" + (RowNumber + 1) + "- F" + (RowNumber + 1) + ")/F" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":G" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;

               sSheet.Set_AllBorders("A1:G" + (RowNumber), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_OutsideBorders("A1", Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("B1:G2", Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_Font_Size("A1:G" + (RowNumber + 1), 12, SheetNumber);
               sSheet.Auto_fit(0, 6, SheetNumber);

               /**************************************************************************************************************************
               * YTD LOWER THAN LAST YEAR
               *************************************************************************************************************************/

               SheetNumber++;
               RowNumber = 0;

               sSheet.Insert_Worksheet("YTD BELOW LAST YEAR", SheetNumber);

               sSheet.Set_Cell(RowNumber, 1, "TOTALS", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Merge_Cells("B1:D1", SheetNumber);
               sSheet.Set_Cell(RowNumber, 4, "TONNAGE TOTALS", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Merge_Cells("E1:G1", SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 0, "THIS YR V LAST YR", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Rotation("A2", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
               sSheet.Set_Font_Size("A2", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
               sSheet.Merge_Cells("A1:A2", SheetNumber);

               sSheet.Set_Cell(RowNumber, 1, Year + " YTD", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 2, LastYear + " YTD", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 3, "VAR AGAINST " + LastYear, SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 4, Year + " YTD TONNAGE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 5, LastYear + " YTD TONNAGE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 6, "VAR AGAINST " + LastYear, SheetNumber, SpreadsheetHorizontalAlignment.Center);

               sSheet.Set_FontColour("A1:A2", LightGreen, Color.Black, SheetNumber);
               sSheet.Set_FontColour("B1:G2", Color.LightGray, Color.Black, SheetNumber);

               sSheet.Set_Bold_Range("A1:G2", true, SheetNumber);

               RowNumber++;

               foreach (DataRow row in YTDSales.Rows)
               {
                  bool Outdated = OutdatedList.Contains(Classes.Global.ConvertToString(row["Name"]).Trim());
                  bool inBudget = BudgetNameList.Contains(Classes.Global.ConvertToString(row["Name"]).Trim());
                  if (!Outdated || inBudget)
                  {
                     if (Classes.Global.ConvertToString(row["Name"]) != "")
                     {
                        if (Classes.Global.ConvertToString(row["Name"]) != "STANNAH STAIRLIFT EURO ACCOUNT")
                        {
                           sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]).Trim(), SheetNumber);
                           sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:BZ,38,FALSE),0)", SheetNumber, "£ #,##0.00");
                           sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:BZ,41,FALSE),0)", SheetNumber, "£ #,##0.00");
                           sSheet.Set_Formula(RowNumber, 3, "=IFERROR((B" + (RowNumber + 1) + "- C" + (RowNumber + 1) + ")/C" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");
                           sSheet.Set_Formula(RowNumber, 4, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:BZ,39,FALSE),0)", SheetNumber, "#,##0.00");
                           sSheet.Set_Formula(RowNumber, 5, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:BZ,42,FALSE),0)", SheetNumber, "#,##0.00");
                           sSheet.Set_Formula(RowNumber, 6 , "=IFERROR((E" + (RowNumber + 1) + "- F" + (RowNumber + 1) + ")/F" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");

                           double YTDValue = Classes.Global.ConvertToDouble(sSheet.Get_Cell_Text_From_Formula(sSheet.Get_Row_Index_From_Name(Classes.Global.ConvertToString(row["Name"]), Year + " MONTH SALES PER CUSTOMER"), 37, Year + " MONTH SALES PER CUSTOMER").ToString().Replace("£", ""));
                           double LastYTDValue = Classes.Global.ConvertToDouble(sSheet.Get_Cell_Text_From_Formula(sSheet.Get_Row_Index_From_Name(Classes.Global.ConvertToString(row["Name"]), LastYear + " MONTH SALES PER CUSTOMER"), 40, LastYear + " MONTH SALES PER CUSTOMER").ToString().Replace("£", ""));

                           if (YTDValue < LastYTDValue)
                              RowNumber++;
                        }
                     }
                  }
               }

               sSheet.Set_Cell(RowNumber, 0, "TOTAL:", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 1, "=SUBTOTAL(9, B3:B" + RowNumber + ")", SheetNumber, "£ #,##0.00");
               sSheet.Set_Formula(RowNumber, 2, "=SUBTOTAL(9, C3:C" + RowNumber + ")", SheetNumber, "£ #,##0.00");
               sSheet.Set_Formula(RowNumber, 3, "=IFERROR((B" + (RowNumber + 1) + "- C" + (RowNumber + 1) + ")/C" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");
               sSheet.Set_Formula(RowNumber, 4, "=SUBTOTAL(9, E3:E" + RowNumber + ")", SheetNumber, "#,##0.00");
               sSheet.Set_Formula(RowNumber, 5, "=SUBTOTAL(9, F3:F" + RowNumber + ")", SheetNumber, "#,##0.00");
               sSheet.Set_Formula(RowNumber, 6, "=IFERROR((E" + (RowNumber + 1) + "- F" + (RowNumber + 1) + ")/F" + (RowNumber + 1) + ",0)", SheetNumber, "#,##0 %");

               sSheet.Set_Bold_Range("A" + (RowNumber + 1) + ":D" + (RowNumber + 1), true, SheetNumber);

               RowNumber++;

               sSheet.Set_AllBorders("A1:G" + (RowNumber), Color.Black, BorderLineStyle.Thin, SheetNumber);
               sSheet.Set_OutsideBorders("A1", Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_OutsideBorders("B1:G2", Color.Black, SheetNumber, BorderLineStyle.Medium);
               sSheet.Set_Font_Size("A1:G" + (RowNumber + 1), 12, SheetNumber);
               sSheet.Auto_fit(0, 6, SheetNumber);

               /**************************************************************************************************************************
               * TONNAGE THIS YEAR V LAST YEAR
               *************************************************************************************************************************/

               SheetNumber++;
               RowNumber = 0;

               sSheet.Insert_Worksheet("TONNAGE " + Year + " V " + LastYear, SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, "TOP 20", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 16, "STANNAH", SheetNumber, SpreadsheetHorizontalAlignment.Center);

               sSheet.Merge_Cells("A1:O1", SheetNumber);
               sSheet.Merge_Cells("P1:AD1", SheetNumber);
               sSheet.Set_Font_Size("A1:AD1", 24, SheetNumber);

               sSheet.Set_Cell(RowNumber, 30, "YTD TONNAGE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Merge_Cells("AE1:AG1", SheetNumber);
               sSheet.Set_Font_Size("AE1:AG1", 24, SheetNumber);

               sSheet.Set_Bold_Range("A1:AG1", true, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 30, "Name", SheetNumber);
               sSheet.Set_Cell(RowNumber, 31, Year, SheetNumber);
               sSheet.Set_Cell(RowNumber, 32, LastYear, SheetNumber);

               RowNumber++;

               ThisYearSales = ThisYearSales.OrderBy(o => o.Name).ToList();

               List<PriorYearSalesModel> groupedmodel = new List<PriorYearSalesModel>();

               double? TotalWeight = 0;
               string CustomerName = "";

               foreach (PriorYearSalesModel model in ThisYearSales)
               {
                  if (CustomerName != model.Name)
                  {
                     if (CustomerName != "")
                     {
                        PriorYearSalesModel newmodel = new PriorYearSalesModel
                        {
                           Name = CustomerName,
                           LineUnitWeight = TotalWeight
                        };

                        groupedmodel.Add(newmodel);
                     }

                     CustomerName = model.Name;
                     if (model.LineUnitWeight.HasValue)
                        TotalWeight = model.LineUnitWeight;
                     else
                        TotalWeight = 0;
                  }
                  else
                  {
                     CustomerName = model.Name;
                     if (model.LineUnitWeight.HasValue)
                        TotalWeight += model.LineUnitWeight;
                  }
               }

               if (TotalWeight.HasValue)
               {
                  PriorYearSalesModel newmodel = new PriorYearSalesModel
                  {
                     Name = CustomerName,
                     LineUnitWeight = TotalWeight.Value
                  };

                  groupedmodel.Add(newmodel);
               }

               groupedmodel = groupedmodel.OrderByDescending(o => o.LineUnitWeight).ToList();

               for (int i = 0; i < 21; i++)
               {
                  if (Classes.Global.ConvertToString(groupedmodel[i].Name).Trim() == "STANNAH STAIRLIFT EURO ACCOUNT")
                     sSheet.Set_Cell(RowNumber, 30, "STANNAH STAIRLIFTS LTD", SheetNumber);
                  else
                     sSheet.Set_Cell(RowNumber, 30, Classes.Global.ConvertToString(groupedmodel[i].Name).Trim(), SheetNumber);

                  sSheet.Set_Formula(RowNumber, 31, "=IFERROR(VLOOKUP(AE" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,42,FALSE),0)", SheetNumber, "#,##0");
                  sSheet.Set_Formula(RowNumber, 32, "=IFERROR(VLOOKUP(AE" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AZ,42,FALSE),0)", SheetNumber, "#,##0");

                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 30, "TOTAL W/O STANNAH", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 31, "=SUBTOTAL(109,AF4:AF" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 32, "=SUBTOTAL(109,AG4:AG" + RowNumber + ")", SheetNumber, "#,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 30, Month + " TONNAGE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Merge_Cells("AE25:AG25", SheetNumber);
               sSheet.Set_Font_Size("AE25:AG25", 24, SheetNumber);
               sSheet.Set_Bold_Range("AE25:AG25", true, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 30, "Name", SheetNumber);
               sSheet.Set_Cell(RowNumber, 31, Month + " " + Year, SheetNumber);
               sSheet.Set_Cell(RowNumber, 32, Month + " " + LastYear, SheetNumber);

               RowNumber++;

               groupedmodel = new List<PriorYearSalesModel>();

               TotalWeight = 0;
               CustomerName = "";

               foreach (PriorYearSalesModel model in ThisYearSales.Where(w => w.InvoiceMonth == (MonthColumnIndex - 1)))
               {
                  if (CustomerName != model.Name)
                  {
                     if (CustomerName != "")
                     {
                        PriorYearSalesModel newmodel = new PriorYearSalesModel
                        {
                           Name = CustomerName,
                           LineUnitWeight = TotalWeight
                        };

                        groupedmodel.Add(newmodel);
                     }

                     CustomerName = model.Name;
                     if (model.LineUnitWeight.HasValue)
                        TotalWeight = model.LineUnitWeight;
                     else
                        TotalWeight = 0;
                  }
                  else
                  {
                     CustomerName = model.Name;
                     if (model.LineUnitWeight.HasValue)
                        TotalWeight += model.LineUnitWeight;
                  }
               }

               if (TotalWeight.HasValue)
               {
                  PriorYearSalesModel newmodel = new PriorYearSalesModel
                  {
                     Name = CustomerName,
                     LineUnitWeight = TotalWeight.Value
                  };

                  groupedmodel.Add(newmodel);
               }

               groupedmodel = groupedmodel.OrderByDescending(o => o.LineUnitWeight).ToList();

               for (int i = 0; i < 21; i++)
               {
                  if (Classes.Global.ConvertToString(groupedmodel[i].Name).Trim() == "STANNAH STAIRLIFT EURO ACCOUNT")
                     sSheet.Set_Cell(RowNumber, 30, "STANNAH STAIRLIFTS LTD", SheetNumber);
                  else
                     sSheet.Set_Cell(RowNumber, 30, Classes.Global.ConvertToString(groupedmodel[i].Name).Trim(), SheetNumber);

                  sSheet.Set_Formula(RowNumber, 31, "=IFERROR(VLOOKUP(AE" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3)) + ",FALSE),0)", SheetNumber, "#,##0");
                  sSheet.Set_Formula(RowNumber, 32, "=IFERROR(VLOOKUP(AE" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3)) + ",FALSE),0)", SheetNumber, "#,##0");

                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 30, "TOTAL W/O STANNAH", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 31, "=SUBTOTAL(109,AF28:AF" + RowNumber + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 32, "=SUBTOTAL(109,AG28:AG" + RowNumber + ")", SheetNumber, "#,##0");

               Color VisionBlue = Color.FromArgb(5, 178, 179);
               Color VisionGray = Color.FromArgb(141, 142, 144);

               sms.Set_Chart(sSheet, "AF4:AF23", "AG4:AG23", "AF2", "AG2", "AE4:AE23", "AE4:AE23", "A2", "O24", SheetNumber, ChartType.ColumnClustered, Color.MediumSeaGreen, VisionGray, Color.MediumSeaGreen, VisionGray, Title: "YTD TOP 20 TONNAGE " + Year + " V " + LastYear + "");
               sms.Set_Chart(sSheet, "AF3:AF3", "AG3:AG3", "AF2", "AG2", "AE3:AE3", "AE3:AE3", "P2", "AD24", SheetNumber, ChartType.ColumnClustered, Color.MediumSeaGreen, VisionGray, Color.MediumSeaGreen, VisionGray, Title: "YTD STANNAH TONNAGE " + Year + " V " + LastYear + "");
               sms.Set_Chart(sSheet, "AF28:AF47", "AG28:AG47", "AF26", "AG26", "AE28:AE47", "AE28:AE47", "A26", "O48", SheetNumber, ChartType.ColumnClustered, VisionBlue, VisionGray, VisionBlue, VisionGray, Title: Month + " TOP 20 TONNAGE " + Year + " V " + LastYear + "");
               sms.Set_Chart(sSheet, "AF27:AF27", "AG27:AG27", "AF26", "AG26", "AE27:AE27", "AE27:AE27", "P26", "AD48", SheetNumber, ChartType.ColumnClustered, VisionBlue, VisionGray, VisionBlue, VisionGray, Title: Month + " STANNAH TONNAGE " + Year + " V " + LastYear + "");


               sSheet.Auto_fit(30, 32, SheetNumber);


               /**************************************************************************************************************************
               * REVENUE THIS YEAR V LAST YEAR
               *************************************************************************************************************************/

               SheetNumber++;
               RowNumber = 0;

               sSheet.Insert_Worksheet("REVENUE " + Year + " V " + LastYear, SheetNumber);

               sSheet.Set_Cell(RowNumber, 0, "TOP 20", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Set_Cell(RowNumber, 16, "STANNAH", SheetNumber, SpreadsheetHorizontalAlignment.Center);

               sSheet.Merge_Cells("A1:O1", SheetNumber);
               sSheet.Merge_Cells("P1:AD1", SheetNumber);
               sSheet.Set_Font_Size("A1:AD1", 24, SheetNumber);

               sSheet.Set_Cell(RowNumber, 30, "YTD REVENUE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Merge_Cells("AE1:AG1", SheetNumber);
               sSheet.Set_Font_Size("AE1:AG1", 24, SheetNumber);

               sSheet.Set_Bold_Range("A1:AG1", true, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 30, "Name", SheetNumber);
               sSheet.Set_Cell(RowNumber, 31, Year, SheetNumber);
               sSheet.Set_Cell(RowNumber, 32, LastYear, SheetNumber);

               RowNumber++;

               YTDSales = resort(YTDSales, "Line_Sale_Price", "DESC");

               for (int i = 0; i < 21; i++)
               {
                  if (Classes.Global.ConvertToString(YTDSales.Rows[i]["Name"]).Trim() == "STANNAH STAIRLIFT EURO ACCOUNT")
                     sSheet.Set_Cell(RowNumber, 30, "STANNAH STAIRLIFTS LTD", SheetNumber);
                  else
                     sSheet.Set_Cell(RowNumber, 30, Classes.Global.ConvertToString(YTDSales.Rows[i]["Name"]).Trim(), SheetNumber);

                  sSheet.Set_Formula(RowNumber, 31, "=IFERROR(VLOOKUP(AE" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ,41,FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 32, "=IFERROR(VLOOKUP(AE" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AZ,41,FALSE),0)", SheetNumber, "£ #,##0");

                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 30, "TOTAL W/O STANNAH", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 31, "=SUBTOTAL(109,AF4:AF" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 32, "=SUBTOTAL(109,AG4:AG" + RowNumber + ")", SheetNumber, "£ #,##0");

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 30, Month + " REVENUE", SheetNumber, SpreadsheetHorizontalAlignment.Center);
               sSheet.Merge_Cells("AE25:AG25", SheetNumber);
               sSheet.Set_Font_Size("AE25:AG25", 24, SheetNumber);
               sSheet.Set_Bold_Range("AE25:AG25", true, SheetNumber);

               RowNumber++;

               sSheet.Set_Cell(RowNumber, 30, "Name", SheetNumber);
               sSheet.Set_Cell(RowNumber, 31, Month + " " + Year, SheetNumber);
               sSheet.Set_Cell(RowNumber, 32, Month + " " + LastYear, SheetNumber);

               RowNumber++;

               MonthTurnoverTable = resort(MonthTurnoverTable, "Line_Sale_Price", "DESC");

               for (int i = 0; i < 21; i++)
               {
                  if (Classes.Global.ConvertToString(MonthTurnoverTable.Rows[i]["Name"]).Trim() == "STANNAH STAIRLIFT EURO ACCOUNT")
                     sSheet.Set_Cell(RowNumber, 30, "STANNAH STAIRLIFTS LTD", SheetNumber);
                  else
                     sSheet.Set_Cell(RowNumber, 30, Classes.Global.ConvertToString(MonthTurnoverTable.Rows[i]["Name"]).Trim(), SheetNumber);

                  sSheet.Set_Formula(RowNumber, 31, "=IFERROR(VLOOKUP(AE" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3) - 1) + ",FALSE),0)", SheetNumber, "£ #,##0");
                  sSheet.Set_Formula(RowNumber, 32, "=IFERROR(VLOOKUP(AE" + (RowNumber + 1) + ",'" + LastYear + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3) - 1) + ",FALSE),0)", SheetNumber, "£ #,##0");

                  RowNumber++;
               }

               sSheet.Set_Cell(RowNumber, 30, "TOTAL W/O STANNAH", SheetNumber, SpreadsheetHorizontalAlignment.Right);
               sSheet.Set_Formula(RowNumber, 31, "=SUBTOTAL(109,AF28:AF" + RowNumber + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 32, "=SUBTOTAL(109,AG28:AG" + RowNumber + ")", SheetNumber, "£ #,##0");

               sms.Set_Chart(sSheet, "AF4:AF23", "AG4:AG23", "AF2", "AG2", "AE4:AE23", "AE4:AE23", "A2", "O24", SheetNumber, ChartType.ColumnClustered, Color.MediumBlue, VisionGray, Color.MediumBlue, VisionGray, Title: "YTD TOP 20 REVENUE " + Year + " V " + LastYear + "");
               sms.Set_Chart(sSheet, "AF3:AF3", "AG3:AG3", "AF2", "AG2", "AE3:AE3", "AE3:AE3", "P2", "AD24", SheetNumber, ChartType.ColumnClustered, Color.MediumBlue, VisionGray, Color.MediumBlue, VisionGray, Title: "YTD STANNAH REVENUE " + Year + " V " + LastYear + "");
               sms.Set_Chart(sSheet, "AF28:AF47", "AG28:AG47", "AF26", "AG26", "AE28:AE47", "AE28:AE47", "A26", "O48", SheetNumber, ChartType.ColumnClustered, Color.Orange, VisionGray, Color.Orange, VisionGray, Title: Month + " TOP 20 REVENUE " + Year + " V " + LastYear + "");
               sms.Set_Chart(sSheet, "AF27:AF27", "AG27:AG27", "AF26", "AG26", "AE27:AE27", "AE27:AE27", "P26", "AD48", SheetNumber, ChartType.ColumnClustered, Color.Orange, VisionGray, Color.Orange, VisionGray, Title: Month + " STANNAH REVENUE " + Year + " V " + LastYear + "");


               sSheet.Auto_fit(30, 32, SheetNumber);

            }
            catch (Exception ex)
            {
               if (Classes.Global.hasArgs)
                  Create_Error_Log(ModuleName, "cmdGo_Click", ex.Message, ex);
               else
                  DXTools.ProcessError.Show(ModuleName, "cmdGo_Click", ex);
            }
            finally
            {
               /*************************************************************************************************************************
              * CHANGE SHEET ORDER
              *************************************************************************************************************************/
               string Year = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).ToString("yyyy");
               string LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy");
               string PriorYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-2).ToString("yyyy");
               string Month = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).ToString("MMMM").ToUpper();

               sSheet.Change_Worksheet_Order("1.TOP 15", 0);
               sSheet.Change_Worksheet_Order("TONNAGE " + Year + " V " + LastYear, 1);
               sSheet.Change_Worksheet_Order("REVENUE " + Year + " V " + LastYear, 2);
               sSheet.Change_Worksheet_Order("2.NEW B V BUDGET", 3);
               sSheet.Change_Worksheet_Order("3.SALES V PRIOR YEARS", 4);
               sSheet.Change_Worksheet_Order("4.CUSTOMER £ v BDG v PY", 5);
               sSheet.Change_Worksheet_Order("5." + Month + " SALES £&KG", 6);
               sSheet.Change_Worksheet_Order("6.SG INTERNAL SALES", 7);
               sSheet.Change_Worksheet_Order(Year + " MONTH SALES PER CUSTOMER", 8);
               sSheet.Change_Worksheet_Order(Year + " BUDGET", 9);
               sSheet.Change_Worksheet_Order("FORECAST WITH WEIGHT", 10);
               sSheet.Change_Worksheet_Order("YTD BELOW BUDGET", 11);
               sSheet.Change_Worksheet_Order("YTD BELOW LAST YEAR", 12);
               sSheet.Change_Worksheet_Order("SALES V BUDGET", 13);
               sSheet.Change_Worksheet_Order("YTD SALES", 14);
               sSheet.Change_Worksheet_Order("CURRENT MONTH TURNOVER SUMMARY", 15);
               sSheet.Change_Worksheet_Order(LastYear + " MONTH SALES PER CUSTOMER", 16);
               sSheet.Change_Worksheet_Order(PriorYear + " MONTH SALES PER CUSTOMER", 17);
               sSheet.Change_Worksheet_Order(LastYear + " BUDGET", 18);
               sSheet.Change_Worksheet_Order(PriorYear + " BUDGET", 19);

               sSheet.Change_Worksheet_Colour(0, LightGreen);
               sSheet.Change_Worksheet_Colour(1, LightGreen);
               sSheet.Change_Worksheet_Colour(2, LightGreen);
               sSheet.Change_Worksheet_Colour(3, LightGreen);
               sSheet.Change_Worksheet_Colour(4, LightGreen);
               sSheet.Change_Worksheet_Colour(5, LightGreen);
               sSheet.Change_Worksheet_Colour(6, LightGreen);
               sSheet.Change_Worksheet_Colour(7, LightGreen);
               sSheet.Change_Worksheet_Colour(8, LightGreen);
               sSheet.Change_Worksheet_Colour(9, Color.LightBlue);
               sSheet.Change_Worksheet_Colour(10, Color.LightBlue);
               sSheet.Change_Worksheet_Colour(11, Color.LightBlue);
               sSheet.Change_Worksheet_Colour(12, Color.LightGreen);
               sSheet.Change_Worksheet_Colour(13, Color.LightGreen);
               sSheet.Change_Worksheet_Colour(14, Color.LightGreen);
               sSheet.Change_Worksheet_Colour(15, Color.Red);
               sSheet.Change_Worksheet_Colour(16, Color.Red);
               sSheet.Change_Worksheet_Colour(17, Color.Red);
               sSheet.Change_Worksheet_Colour(18, Color.Red);
               sSheet.Change_Worksheet_Colour(19, Color.Red);

               for (int i = 20; i < sSheet.Get_Sheet_Count(); i++)
                  sSheet.Change_Worksheet_Colour(i, Color.Gold);


               /**************************************************************************************************************************
               * FINAL BITS AND SAVING
               *************************************************************************************************************************/

               sSheet.Set_Active_Sheet(0);

               if (!Classes.Global.hasArgs)
               {
                  sSheet.Hide_Wait();

                  // Save the sheet 
                  using (SaveFileDialog saveDialog = new SaveFileDialog())
                  {
                     saveDialog.Filter = "Excel (2010) (.xlsx)|*.xlsx";
                     saveDialog.Title = "Save Spreadsheet";
                     saveDialog.FileName = "Monthly Sales Analysis " + System.DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
                     if (saveDialog.ShowDialog() != DialogResult.Cancel)
                     {
                        sSheet.SaveToFile(saveDialog.FileName, DXTools.Spreadsheet.FormatTypes.Xlsx);
                        System.Diagnostics.Process.Start(saveDialog.FileName);
                     }


                  }
               }
               else
               {
                  clsConfig cFig = new clsConfig();
                  cFig.Retrieve("Where ConfigID = 'Monthly Sales Analysis Email Dist'");

                  if (cFig.cValue != null)
                     new Classes.SupportMail().Send("Please find the attached Monthly Sales Analysis Sheet for " + Month + " " + Year, "Monthly Sales Analysis", cFig.cValue, new List<Stream>() { sSheet.Export_To_Stream(DXTools.Spreadsheet.FormatTypes.Xlsx) }, "Monthly Sales Analysis " + Month + " " + Year + ".xlsx");
                  else
                     new Classes.SupportMail().Send("Please find the attached Monthly Sales Analysis Sheet for " + Month + " " + Year, "Monthly Sales Analysis", "nathan.staples@visionprofiles.co.uk", new List<Stream>() { sSheet.Export_To_Stream(DXTools.Spreadsheet.FormatTypes.Xlsx) }, "Monthly Sales Analysis " + Month + " " + Year + ".xlsx");
               }
            }
         }
      }



      private List<PriorYearSalesModel> MonthSalesPerCustomerSheets(Spreadsheet sSheet, int SheetNumber, string Year, string EndDate, Color LightGreen, int MonthNo = 12)
      {
         try
         {
            /**************************************************************************************************************************
            * PRIOR / THIS YEAR MONTH SALES PER CUSTOMER
            **************************************************************************************************************************/

            int RowNumber = 0;
            clsInvoices Invoices = new clsInvoices();

            sSheet.Add_Worksheet(Year + " MONTH SALES PER CUSTOMER");

            string sqlstring = "SELECT tbl_Customer.CustomerID, tbl_Customer.Account_Ref, LTRIM(RTRIM(tbl_Customer.Name)) AS Name, Inv.Line_Cost_Price, " +
                    "Inv.Line_Sale_Price, Inv.Line_Unit_Weight, Inv.Invoice_Month, tbl_Customer.Deleted " +
                    "FROM tbl_Customer LEFT OUTER JOIN(SELECT SUM(tbl_InvoiceItem.Cost_Price* tbl_InvoiceItem.Qty_Order) AS Line_Cost_Price, SUM(tbl_InvoiceItem.Net_Amount) AS Line_Sale_Price, " +
                    "SUM(tbl_Product.Unit_Weight_Without_Components * tbl_InvoiceItem.Qty_Order) AS Line_Unit_Weight, " +
                    "tbl_Invoice.CustomerID, MONTH(tbl_Invoice.Invoice_Date) AS Invoice_Month " +
                    "FROM tbl_Invoice AS tbl_Invoice LEFT OUTER JOIN " +
                    "tbl_Product RIGHT OUTER JOIN " +
                    "tbl_InvoiceItem ON tbl_Product.ProductID = tbl_InvoiceItem.ProductID ON tbl_Invoice.InvoiceID = tbl_InvoiceItem.InvoiceID " +
                    "WHERE(tbl_Invoice.Invoice_Date IS NULL OR " +
                    "tbl_Invoice.Invoice_Date BETWEEN'" + Year + "-01-01 00:00:00' AND '" + EndDate + "') " +
                    "GROUP BY tbl_Invoice.CustomerID, MONTH(Invoice_Date)) Inv ON tbl_Customer.CustomerID = Inv.CustomerID " +
            "WHERE (tbl_Customer.Deleted = 0 OR Inv.Line_Sale_Price > 0)" +
                    "ORDER BY tbl_Customer.Name ";

            DataTable PriorYearSalesTable = Invoices.RetrieveDataTable(sqlstring, false);



            List<PriorYearSalesModel> PriorYearSaleList = PriorYearSalesTable.AsEnumerable().Select(s => new PriorYearSalesModel
            {
               AccountRef = s.Field<string>("Account_Ref"),
               Name = s.Field<string>("Name"),
               LineCostPrice = s.Field<double?>("Line_Cost_Price"),
               LineSalePrice = s.Field<double?>("Line_Sale_Price"),
               LineUnitWeight = s.Field<double?>("Line_Unit_Weight"),
               InvoiceMonth = s.Field<int?>("Invoice_Month"),
               Deleted = s.Field<bool>("Deleted")
            }).ToList();

            // Merge Stannah Accounts & New Wave Accounts

            for (int i = 1; i <= 12; i++)
            {
               PriorYearSalesModel StannahLtd = PriorYearSaleList.Where(w => w.InvoiceMonth == i && w.Name == "STANNAH STAIRLIFTS LTD").FirstOrDefault();
               PriorYearSalesModel StannahEuro = PriorYearSaleList.Where(w => w.InvoiceMonth == i && w.Name == "STANNAH STAIRLIFT EURO ACCOUNT").FirstOrDefault();

               if (StannahLtd != null && StannahEuro != null)
               {
                  StannahLtd.LineCostPrice += StannahEuro.LineCostPrice;
                  StannahLtd.LineSalePrice += StannahEuro.LineSalePrice;
                  StannahLtd.LineUnitWeight += StannahEuro.LineUnitWeight;

                  PriorYearSaleList.Remove(StannahEuro);
               }
               else if (StannahLtd == null && StannahEuro != null)
               {
                  StannahEuro.Name = "STANNAH STAIRLIFTS LTD";
                  StannahEuro.Deleted = false;
               }

               PriorYearSalesModel NewWaveDirect = PriorYearSaleList.Where(w => w.InvoiceMonth == i && w.Name == "NEW WAVE DOORS DIRECT LTD").FirstOrDefault();
               PriorYearSalesModel Deltaco = PriorYearSaleList.Where(w => w.InvoiceMonth == i && w.Name == "DELTACO 1 LTD").FirstOrDefault();

               if (NewWaveDirect != null && Deltaco != null)
               {
                  NewWaveDirect.LineCostPrice += Deltaco.LineCostPrice;
                  NewWaveDirect.LineSalePrice += Deltaco.LineSalePrice;
                  NewWaveDirect.LineUnitWeight += Deltaco.LineUnitWeight;

                  PriorYearSaleList.Remove(Deltaco);
               }
               else if (NewWaveDirect == null && Deltaco != null)
               {
                  Deltaco.Name = "NEW WAVE DOORS DIRECT LTD";
                  Deltaco.Deleted = false;
               }

               PriorYearSalesModel Digico = PriorYearSaleList.Where(w => w.InvoiceMonth == i && w.Name == "DIGICO (UK) LTD").FirstOrDefault();
               PriorYearSalesModel A6Audio = PriorYearSaleList.Where(w => w.InvoiceMonth == i && w.Name == "A6 AUDIO LIMITED").FirstOrDefault();

               if (Digico != null && A6Audio != null)
               {
                  Digico.LineCostPrice += A6Audio.LineCostPrice;
                  Digico.LineSalePrice += A6Audio.LineSalePrice;
                  Digico.LineUnitWeight += A6Audio.LineUnitWeight;

                  PriorYearSaleList.Remove(A6Audio);
               }
               else if (Digico == null && A6Audio != null)
               {
                  A6Audio.Name = "DIGICO (UK) LTD";
                  A6Audio.Deleted = false;
               }
            }

            PriorYearSalesTable.Dispose();
            PriorYearSalesTable = null;

            sSheet.Set_Cell(RowNumber, 1, "JAN", SheetNumber);
            sSheet.Set_Cell(RowNumber, 4, "FEB", SheetNumber);
            sSheet.Set_Cell(RowNumber, 7, "MAR", SheetNumber);
            sSheet.Set_Cell(RowNumber, 10, "APR", SheetNumber);
            sSheet.Set_Cell(RowNumber, 13, "MAY", SheetNumber);
            sSheet.Set_Cell(RowNumber, 16, "JUN", SheetNumber);
            sSheet.Set_Cell(RowNumber, 19, "JUL", SheetNumber);
            sSheet.Set_Cell(RowNumber, 22, "AUG", SheetNumber);
            sSheet.Set_Cell(RowNumber, 25, "SEP", SheetNumber);
            sSheet.Set_Cell(RowNumber, 28, "OCT", SheetNumber);
            sSheet.Set_Cell(RowNumber, 31, "NOV", SheetNumber);
            sSheet.Set_Cell(RowNumber, 34, "DEC", SheetNumber);
            sSheet.Set_Cell(RowNumber, 37, "ANNUAL TOTAL", SheetNumber);
            sSheet.Set_Cell(RowNumber, 40, "YTD TOTAL", SheetNumber);

            sSheet.Set_Cell(RowNumber, 0, Year, SheetNumber, SpreadsheetHorizontalAlignment.Center);
            sSheet.Set_Rotation("A1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
            sSheet.Set_Font_Size("A1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
            sSheet.Set_Bold_Range("A1:AQ1", true, SheetNumber);
            sSheet.Merge_Cells("A1:A2", SheetNumber);

            sSheet.Set_FontColour("A1:AK2", LightGreen, Color.Black, SheetNumber);
            sSheet.Set_FontColour("B1:AN1", Color.LightGray, Color.Black, SheetNumber);
            RowNumber++;


            for (int i = 1; i < 40; i += 3)
            {
               sSheet.Set_Cell(RowNumber, i, "VALUE", SheetNumber);
               sSheet.Set_Cell(RowNumber, i + 1, "WEIGHT", SheetNumber);
               sSheet.Set_Cell(RowNumber, i + 2, "ASP", SheetNumber);
            }

            sSheet.Set_Cell(RowNumber, 40, "YTD VALUE", SheetNumber);
            sSheet.Set_Cell(RowNumber, 41, "YTD WEIGHT", SheetNumber);
            sSheet.Set_Cell(RowNumber, 42, "YTD ASP", SheetNumber);

            RowNumber++;

            sSheet.FormatCell("B:AQ", "£ #,##0", SheetNumber);

            PriorYearSaleList = PriorYearSaleList.OrderBy(o => o.Name).ToList();

            foreach (var priorYearSalesGroup in PriorYearSaleList.GroupBy(g => new { g.Name, g.Deleted }))
            {
               sSheet.Set_Cell(RowNumber, 0, priorYearSalesGroup.Key.Name, SheetNumber);

               if (priorYearSalesGroup.Key.Deleted)
               {
                  if (priorYearSalesGroup.Key.Name != "NEW WAVE DOORS DIRECT LTD")
                     sSheet.Set_FontColour("A" + (RowNumber + 1) + ":AK" + (RowNumber + 1), Color.Gray, Color.White, SheetNumber);
               }

               foreach (PriorYearSalesModel priorYearSale in PriorYearSaleList.Where(w => w.Name == priorYearSalesGroup.Key.Name))
               {
                  if (priorYearSale.InvoiceMonth.HasValue)
                  {
                     int StartColumn = (priorYearSale.InvoiceMonth.Value * 3) - 2;

                     sSheet.Set_Cell(RowNumber, (StartColumn), priorYearSale.LineSalePrice, SheetNumber, SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Cell(RowNumber, (StartColumn + 1), priorYearSale.LineUnitWeight, SheetNumber, SpreadsheetHorizontalAlignment.Right);
                     sSheet.Set_Formula(RowNumber, (StartColumn + 2), "=IFERROR(" + sSheet.GetExcelColumnName(StartColumn + 1) + (RowNumber + 1) + "/" + sSheet.GetExcelColumnName(StartColumn + 2) + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");
                  }
               }

               sSheet.Set_Formula(RowNumber, 37, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + "+Q" + (RowNumber + 1) + "+T" + (RowNumber + 1) + "+W" + (RowNumber + 1) +
               "+Z" + (RowNumber + 1) + "+AC" + (RowNumber + 1) + "+AF" + (RowNumber + 1) + "+AI" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
               sSheet.Set_Formula(RowNumber, 38, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + "+I" + (RowNumber + 1) + "+L" + (RowNumber + 1) + "+O" + (RowNumber + 1) + "+R" + (RowNumber + 1) + "+U" + (RowNumber + 1) + "+X" + (RowNumber + 1) +
               "+AA" + (RowNumber + 1) + "+AD" + (RowNumber + 1) + "+AG" + (RowNumber + 1) + "+AJ" + (RowNumber + 1) + ")", SheetNumber, "#,##0");
               sSheet.Set_Formula(RowNumber, 39, "=IFERROR(AL" + (RowNumber + 1) + "/AM" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");
               switch (MonthNo)
               {
                  case 1:
                     sSheet.Set_Formula(RowNumber, 40, "=IFERROR(B" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");
                     sSheet.Set_Formula(RowNumber, 41, "=IFERROR(C" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");
                     break;
                  case 2:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 3:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + "+I" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 4:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + "+I" + (RowNumber + 1) + "+L" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 5:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + "+I" + (RowNumber + 1) + "+L" + (RowNumber + 1) + "+O" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 6:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + "+Q" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + "+I" + (RowNumber + 1) + "+L" + (RowNumber + 1) + "+O" + (RowNumber + 1) + "+R" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 7:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + "+Q" + (RowNumber + 1) + "+T" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + "+I" + (RowNumber + 1) + "+L" + (RowNumber + 1) + "+O" + (RowNumber + 1) + "+R" + (RowNumber + 1) + "+U" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 8:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + "+Q" + (RowNumber + 1) + "+T" + (RowNumber + 1) + "+W" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + "+I" + (RowNumber + 1) + "+L" + (RowNumber + 1) + "+O" + (RowNumber + 1) + "+R" + (RowNumber + 1) + "+U" + (RowNumber + 1) + "+X" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 9:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + "+Q" + (RowNumber + 1) + "+T" + (RowNumber + 1) + "+W" + (RowNumber + 1) +
             "+Z" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + "+F" + (RowNumber + 1) + "+I" + (RowNumber + 1) + "+L" + (RowNumber + 1) + "+O" + (RowNumber + 1) + "+R" + (RowNumber + 1) + "+U" + (RowNumber + 1) + "+X" + (RowNumber + 1) +
             "+AA" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 10:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + "+Q" + (RowNumber + 1) + "+T" + (RowNumber + 1) + "+W" + (RowNumber + 1) +
             "+Z" + (RowNumber + 1) + "+AC" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + " + F" + (RowNumber + 1) + " + I" + (RowNumber + 1) + " + L" + (RowNumber + 1) + " + O" + (RowNumber + 1) + " + R" + (RowNumber + 1) + " + U" + (RowNumber + 1) + " + X" + (RowNumber + 1) +
             "+AA" + (RowNumber + 1) + "+AD" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 11:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + "+Q" + (RowNumber + 1) + "+T" + (RowNumber + 1) + "+W" + (RowNumber + 1) +
             "+Z" + (RowNumber + 1) + "+AC" + (RowNumber + 1) + "+AF" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + " + F" + (RowNumber + 1) + " + I" + (RowNumber + 1) + " + L" + (RowNumber + 1) + " + O" + (RowNumber + 1) + " + R" + (RowNumber + 1) + " + U" + (RowNumber + 1) + " + X" + (RowNumber + 1) +
             "+AA" + (RowNumber + 1) + "+AD" + (RowNumber + 1) + "+AG" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
                  case 12:
                     sSheet.Set_Formula(RowNumber, 40, "=SUM(B" + (RowNumber + 1) + "+E" + (RowNumber + 1) + "+H" + (RowNumber + 1) + "+K" + (RowNumber + 1) + "+N" + (RowNumber + 1) + "+Q" + (RowNumber + 1) + "+T" + (RowNumber + 1) + "+W" + (RowNumber + 1) +
             "+Z" + (RowNumber + 1) + "+AC" + (RowNumber + 1) + "+AF" + (RowNumber + 1) + "+AI" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     sSheet.Set_Formula(RowNumber, 41, "=SUM(C" + (RowNumber + 1) + " + F" + (RowNumber + 1) + " + I" + (RowNumber + 1) + " + L" + (RowNumber + 1) + " + O" + (RowNumber + 1) + " + R" + (RowNumber + 1) + " + U" + (RowNumber + 1) + " + X" + (RowNumber + 1) +
             "+AA" + (RowNumber + 1) + "+AD" + (RowNumber + 1) + "+AG" + (RowNumber + 1) + "+AJ" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0");
                     break;
               }
               sSheet.Set_Formula(RowNumber, 42, "=IFERROR(AO" + (RowNumber + 1) + "/AP" + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

               RowNumber++;
            }

            sSheet.Set_Column_Width(0, 46.43, Year + " MONTH SALES PER CUSTOMER");

            for (int i = 1; i <= 42; i++)
               sSheet.Set_Column_Width(i, 11.86, Year + " MONTH SALES PER CUSTOMER");

            sSheet.Set_FontColour("AL1:AQ" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);
            sSheet.Set_AllBorders("A1:AQ" + RowNumber, Color.Black, BorderLineStyle.Thin, SheetNumber);
            sSheet.Set_OutsideBorders("A1:A" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
            sSheet.Set_OutsideBorders("E1:G" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
            sSheet.Set_OutsideBorders("K1:M" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
            sSheet.Set_OutsideBorders("Q1:S" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
            sSheet.Set_OutsideBorders("W1:Y" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
            sSheet.Set_OutsideBorders("AC1:AE" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
            sSheet.Set_OutsideBorders("AI1:AK" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
            sSheet.Set_OutsideBorders("A1:AQ" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);

            sSheet.Set_Cell(RowNumber, 0, "TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Left);

            for (int i = 1; i < 39; i += 3)
            {
               sSheet.Set_Formula(RowNumber, i, "=SUM(" + sSheet.GetExcelColumnName(i + 1) + "3:" + sSheet.GetExcelColumnName(i + 1) + (RowNumber) + ")", SheetNumber, "£ #,##0.00");
               sSheet.Set_Formula(RowNumber, (i + 1), "=SUM(" + sSheet.GetExcelColumnName(i + 2) + "3:" + sSheet.GetExcelColumnName(i + 2) + (RowNumber) + ")", SheetNumber);
               sSheet.Set_Formula(RowNumber, (i + 2), "=IFERROR(" + sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1) + "/" + sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");
            }

            sSheet.Set_Formula(RowNumber, 40, "=SUM(" + sSheet.GetExcelColumnName(41) + "3:" + sSheet.GetExcelColumnName(41) + (RowNumber) + ")", SheetNumber, "£ #,##0.00");
            sSheet.Set_Formula(RowNumber, 41, "=SUM(" + sSheet.GetExcelColumnName(42) + "3:" + sSheet.GetExcelColumnName(42) + (RowNumber) + ")", SheetNumber);
            sSheet.Set_Formula(RowNumber, 42, "=IFERROR(" + sSheet.GetExcelColumnName(41) + (RowNumber + 1) + "/" + sSheet.GetExcelColumnName(42) + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");

            RowNumber += 2;

            sSheet.Set_Tree_Map(SheetNumber, "A" + (RowNumber + 1), "N" + (RowNumber + 41), "YTD Sales Tree Map", "A3:A" + (RowNumber - 2), "AO3:AO" + (RowNumber - 2));

            RowNumber -= 2;
            sSheet.Set_AllBorders("A" + (RowNumber + 1) + ":AQ" + (RowNumber + 1), Color.Black, SheetNumber);
            sSheet.Set_OutsideBorders("A" + (RowNumber + 1) + ":AQ" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
            sSheet.Set_FontColour("A" + (RowNumber + 1) + ":AQ" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

            sSheet.FormatCell("C:C", "#,##0", SheetNumber);
            sSheet.FormatCell("F:F", "#,##0", SheetNumber);
            sSheet.FormatCell("I:I", "#,##0", SheetNumber);
            sSheet.FormatCell("L:L", "#,##0", SheetNumber);
            sSheet.FormatCell("O:O", "#,##0", SheetNumber);
            sSheet.FormatCell("R:R", "#,##0", SheetNumber);
            sSheet.FormatCell("U:U", "#,##0", SheetNumber);
            sSheet.FormatCell("X:X", "#,##0", SheetNumber);
            sSheet.FormatCell("AA:AA", "#,##0", SheetNumber);
            sSheet.FormatCell("AD:AD", "#,##0", SheetNumber);
            sSheet.FormatCell("AG:AG", "#,##0", SheetNumber);
            sSheet.FormatCell("AJ:AJ", "#,##0", SheetNumber);
            sSheet.FormatCell("AM:AM", "#,##0", SheetNumber);
            sSheet.FormatCell("AP:AP", "#,##0", SheetNumber);

            return PriorYearSaleList;
         }
         catch (Exception ex)
         {
            if (Classes.Global.hasArgs)
               Create_Error_Log(ModuleName, "MonthSalesPerCustomerSheets", ex.Message, ex);
            else
               DXTools.ProcessError.Show(ModuleName, "MonthSalesPerCustomerSheets", ex);

            return new List<PriorYearSalesModel>();
         }
      }


      private List<BudgetModel> ConvertBudgetSheetToDataTable(int SheetIndex, int MaxRows, Spreadsheet sSheet)
      {
         try
         {
            // Need to read through each line until we get a blank space and convert into datatable. If we see more than 4 blank spaces weve reached the end of the spreadsheet
            int RowIndex = 0;
            int StartIndex = 1;
            int EndIndex = 0;
            bool newSection = true;

            //DataTable masterTable = new DataTable();
            List<BudgetModel> masterBudget = new List<BudgetModel>();

            for (int i = 0; i < MaxRows; i++)
            {
               string cValue = Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowIndex, 0, SheetIndex));
               if (string.IsNullOrEmpty(cValue))
               {
                  EndIndex = i;
                  if (EndIndex > 1 && newSection == false)
                  {
                     DataTable SectionTable = sSheet.Export_To_Datatable(SheetIndex, sSheet.GetExcelCellRange(StartIndex + 1, 1, EndIndex + 1, 15, SheetIndex));
                     if (SectionTable.Rows.Count > 0)
                     {
                        foreach (DataRow Row in SectionTable.Rows)
                        {
                           BudgetModel newbudget = new BudgetModel
                           {
                              ExistingCustomers = Classes.Global.ConvertToString(Row[SectionTable.Columns[0].ColumnName]),
                              Jan = Classes.Global.ConvertToDouble(Row["Jan"]),
                              Feb = Classes.Global.ConvertToDouble(Row["Feb"]),
                              Mar = Classes.Global.ConvertToDouble(Row["Mar"]),
                              Apr = Classes.Global.ConvertToDouble(Row["Apr"]),
                              May = Classes.Global.ConvertToDouble(Row["May"]),
                              Jun = Classes.Global.ConvertToDouble(Row["Jun"]),
                              Jul = Classes.Global.ConvertToDouble(Row["Jul"]),
                              Aug = Classes.Global.ConvertToDouble(Row["Aug"]),
                              Sep = Classes.Global.ConvertToDouble(Row["Sep"]),
                              Oct = Classes.Global.ConvertToDouble(Row["Oct"]),
                              Nov = Classes.Global.ConvertToDouble(Row["Nov"]),
                              Dec = Classes.Global.ConvertToDouble(Row["Dec"]),
                              YTD = Classes.Global.ConvertToDouble(Row[13]),
                              FullYearBudget = Classes.Global.ConvertToDouble(Row[14]),
                              Section = SectionTable.Columns[0].ColumnName
                           };

                           masterBudget.Add(newbudget);
                        }

                        newSection = true;
                     }
                  }
               }
               else
               {
                  if (newSection)
                     StartIndex = i;

                  newSection = false;
               }
               RowIndex++;
            }
            return masterBudget;
         }
         catch (Exception ex)
         {
            if (Classes.Global.hasArgs)
               Create_Error_Log(ModuleName, "ConvertBudgetSheetToDataTable", ex.Message, ex);
            else
               DXTools.ProcessError.Show(ModuleName, "ConvertBudgetSheetToDataTable", ex);

            return null;
         }
      }

      private static DataTable resort(DataTable dt, string colName, string direction)
      {
         dt.DefaultView.Sort = colName + " " + direction;
         dt = dt.DefaultView.ToTable();
         return dt;
      }

      private void frmMain_Load(object sender, EventArgs e)
      {
         dteReportDate.EditValue = Classes.Global.ConvertToDateTime(System.DateTime.Now.AddMonths(-1).ToString("01/MM/yyyy"));
      }

      private void SetForecastYTD(Spreadsheet sSheet, int MonthColumnIndex, int RowNumber, int i)
      {
         switch (MonthColumnIndex)
         {
            case 2:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 3:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 4:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 5:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 6:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + " + J" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + " + K" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 7:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + " + J" + (i + 1) + " + L" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + " + K" + (i + 1) + " + M" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 8:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + " + J" + (i + 1) + " + L" + (i + 1) + " + N" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + " + K" + (i + 1) + " + M" + (i + 1) + " + O" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 9:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + " + J" + (i + 1) + " + L" + (i + 1) + " + N" + (i + 1) + " + P" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + " + K" + (i + 1) + " + M" + (i + 1) + " + O" + (i + 1) + " + Q" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 10:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + " + J" + (i + 1) + " + L" + (i + 1) + " + N" + (i + 1) + " + P" + (i + 1) + " + R" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + " + K" + (i + 1) + " + M" + (i + 1) + " + O" + (i + 1) + " + Q" + (i + 1) + " + S" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 11:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + " + J" + (i + 1) + " + L" + (i + 1) + " + N" + (i + 1) + " + P" + (i + 1) + " + R" + (i + 1) + " + T" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + " + K" + (i + 1) + " + M" + (i + 1) + " + O" + (i + 1) + " + Q" + (i + 1) + " + S" + (i + 1) + " + U" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 12:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + " + J" + (i + 1) + " + L" + (i + 1) + " + N" + (i + 1) + " + P" + (i + 1) + " + R" + (i + 1) + " + T" + (i + 1) + " + V" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + " + K" + (i + 1) + " + M" + (i + 1) + " + O" + (i + 1) + " + Q" + (i + 1) + " + S" + (i + 1) + " + U" + (i + 1) + " + W" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
            case 13:
               sSheet.Set_Formula(RowNumber, 25, "=SUM(B" + (i + 1) + " + D" + (i + 1) + " + F" + (i + 1) + " + H" + (i + 1) + " + J" + (i + 1) + " + L" + (i + 1) + " + N" + (i + 1) + " + P" + (i + 1) + " + R" + (i + 1) + " + T" + (i + 1) + " + V" + (i + 1) + " + X" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               sSheet.Set_Formula(RowNumber, 26, "=SUM(C" + (i + 1) + " + E" + (i + 1) + " + G" + (i + 1) + " + I" + (i + 1) + " + K" + (i + 1) + " + M" + (i + 1) + " + O" + (i + 1) + " + Q" + (i + 1) + " + S" + (i + 1) + " + U" + (i + 1) + " + W" + (i + 1) + " + Y" + (i + 1) + ")", "FORECAST WITH WEIGHT", "#,##0");
               break;
         }
      }

      internal static void Create_Error_Log(string myClass, string strProcedureName, string ErrorMessage, Exception e)
      {
         string sqlString = string.Empty;
         try
         {
            clsCustomers DataCon = new clsCustomers();
            sqlString = "Insert Into VPSConfig.dbo.tbl_ErrorLog(Application_Name, Module_Name, Procedure_Name, Error_Message, Additional_Info, Last_Updated, Last_Updated_By, Version_Number) Values('BudgetExcelSheets', '" + myClass + "', '" + strProcedureName + "', '" + ErrorMessage + "','','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','', '')";
            DataCon.Execute_SQL(sqlString, SaltireAPI.Global.sqlConnection, "25154253");
         }
         catch (Exception ex)
         {
            // Write it to the hadr drive
            File.WriteAllText(@"\\vision-fp1\StockManagement\Saltire\VPS Logging\Error " + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"), sqlString + Environment.NewLine + Environment.NewLine + ex.GetBaseException().Message);
         }
      }
   }
}