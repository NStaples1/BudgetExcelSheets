using BudgetExcelSheets.Classes;
using DevExpress.Utils.About;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Filtering.Templates;
using DXTools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SaltireAPI;
using DevExpress.XtraSpreadsheet.Model.CopyOperation;
using DevExpress.Spreadsheet;
using System.Drawing.Text;
using DevExpress.Data.Helpers;
using BudgetExcelSheets.Models;
using DevExpress.XtraSpreadsheet.Utils.Trees;
using DevExpress.XtraExport.Implementation;
using DevExpress.XtraSpreadsheet.Import.Xls;
using System.Data.SqlTypes;
using DevExpress.Spreadsheet.Charts;
using DXTools.Classes;

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
            using (DXTools.Spreadsheet sSheet = new DXTools.Spreadsheet())
            {
                try
                {
                    sSheet.Show_Wait();

                    string StartDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).ToString("yyyy-MM-dd");
                    string EndDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddMonths(1).AddSeconds(-1).ToString("yyyy-MM-dd");
                    string Year = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).ToString("yyyy");
                    string LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy");
                    clsInvoices Invoices = new clsInvoices();
                    int RowNumber = 0;
                    int SheetNumber = 0;
                    Color LightGreen = ColorTranslator.FromHtml("#66FFCC");
                    /**************************************************************************************************************************
                    * BUDGET
                    *************************************************************************************************************************/

                    sSheet.LoadFromFile("C:\\Data\\Budget Sheet.xlsx");

                    sSheet.Set_Workbook_Units(Spreadsheet.DocumentUnits.Point);

                    int MonthColumnIndex = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).Month + 1;
                    int BudgetTotalRow = sSheet.GetWorksheetRange(Year + " BUDGET").RowCount;
                    int priorYearBudgetTotalRow = sSheet.GetWorksheetRange(LastYear + " BUDGET").RowCount;
                    int newCustomerNoBudgetTotalRows = sSheet.GetWorksheetRange("NEW BUSINESS NO BUDGET").RowCount;
                    int NewBusinessWonStart = 0;
                    int NewBusinessWonEnd = 0;
                    int newBusinessStart = 0;

                    List<BudgetModel> BudgetList = ConvertBudgetSheetToDataTable(SheetNumber, BudgetTotalRow, sSheet);

                    RowNumber++;

                    for (int i = 1; i < BudgetTotalRow; i++)
                    {
                        if (!string.IsNullOrEmpty(Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 13, 0))))
                            sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (i + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (i + 1) + ")", 0, "#,##0");

                        if (Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 0, 0)) == "NEW BUSINESS WON IN " + LastYear + " IMPACTING " + Year)
                            NewBusinessWonStart = i + 1;
                        if (Classes.Global.ConvertToString(sSheet.Get_Cell_Value(RowNumber, 0, 0)) == "NEW BUSINESS IN " + Year)
                        {
                            NewBusinessWonEnd = i - 1;
                            newBusinessStart = i + 1;
                        }

                        RowNumber++;
                    }

                    SheetNumber++;
                    RowNumber = 0;

                    SheetNumber++;
                    RowNumber = 0;

                    /**************************************************************************************************************************
                    * CURRENT MONTH TURNOVER SUMMARY 
                    * Current month is equal to previous month
                    *************************************************************************************************************************/

                    RowNumber = 0;
                    SheetNumber++;

                    string sqlstring = "SELECT tbl_Invoice.CustomerID, SUM(tbl_InvoiceItem.Cost_Price * tbl_InvoiceItem.Qty_Order) AS Line_Cost_Price, SUM(tbl_InvoiceItem.Net_Amount) AS Line_Sale_Price, " +
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

                    sSheet.Set_Cell(RowNumber, 0, "Customer Name", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 1, "Cost Price", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 2, "Sales Price", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 3, "Profit", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 4, "AVG M/U%", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 5, "AVG P/S%", SheetNumber);

                    RowNumber++;

                    foreach (DataRow row in MonthTurnoverTable.Rows)
                    {
                        double Profit = Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2);
                        sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]).Trim(), SheetNumber);
                        sSheet.Set_Cell(RowNumber, 1, row["Line_Cost_Price"], SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                        sSheet.Set_Cell(RowNumber, 2, row["Line_Sale_Price"], SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                        sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                        sSheet.Set_Cell(RowNumber, 4, Classes.Global.DivideNum(Profit, Classes.Global.ConvertToDouble(row["Line_Cost_Price"]), 4) * 100, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                        sSheet.Set_Cell(RowNumber, 5, Classes.Global.DivideNum(Profit, Classes.Global.ConvertToDouble(row["Line_Sale_Price"]), 4) * 100, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                        RowNumber++;
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
                        "WHERE tbl_Customer.Deleted = 0 " +
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

                    sSheet.Set_Cell(RowNumber, 0, "Customer Name", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 1, "Cost Price", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 2, "Sales Price", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 3, "Profit", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 4, "AVG M/U%", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 5, "AVG P/S%", SheetNumber);

                    RowNumber++;

                    foreach (DataRow row in YTDSales.Rows)
                    {
                        double Profit = Classes.Global.ConvertToDouble(row["Line_Sale_Price"], 2) - Classes.Global.ConvertToDouble(row["Line_Cost_Price"], 2);
                        sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]).Trim(), SheetNumber);
                        sSheet.Set_Cell(RowNumber, 1, row["Line_Cost_Price"], SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                        sSheet.Set_Cell(RowNumber, 2, row["Line_Sale_Price"], SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                        sSheet.Set_Cell(RowNumber, 3, Profit, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                        sSheet.Set_Cell(RowNumber, 4, Classes.Global.DivideNum(Profit, Classes.Global.ConvertToDouble(row["Line_Cost_Price"]), 4) * 100, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);
                        sSheet.Set_Cell(RowNumber, 5, Classes.Global.DivideNum(Profit, Classes.Global.ConvertToDouble(row["Line_Sale_Price"]), 4) * 100, SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Right);

                        RowNumber++;
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

                    SheetNumber = 0;
                    RowNumber = 1;
                    sSheet.Insert_Worksheet("DATA FOR PRESENTATION", SheetNumber);

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

                    sSheet.Set_Cell("H" + (RowNumber + 1), "Top 15 Customers", 0, SpreadsheetHorizontalAlignment.Center);
                    sSheet.Set_Bold(RowNumber, 0, true, SheetNumber);
                    sSheet.Set_Rotation("H" + (RowNumber + 1), 0, 90, SpreadsheetVerticalAlignment.Center);
                    sSheet.Set_Font_Size("H" + (RowNumber + 1), 14, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
                    sSheet.Set_FontColour("H" + (RowNumber + 1) + ":H" + (RowNumber + 1), LightGreen, Color.Black, SheetNumber);

                    DataTable newMonthTable = resort(MonthTurnoverTable, "Line_Sale_Price", "DESC");

                    DataTable newYTDTable = resort(YTDSales, "Line_Sale_Price", "DESC");

                    for (int i = 0; i < 15; i++)
                    {
                        sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(newMonthTable.Rows[i]["Name"]).Trim(), SheetNumber);
                        sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!A:C,3,FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O," + MonthColumnIndex + ",FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 3, "=" + sSheet.GetExcelColumnName(2) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(3) + (RowNumber + 1), SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 4, "=IFERROR(" + sSheet.GetExcelColumnName(4) + (RowNumber + 1) + "/" + sSheet.GetExcelColumnName(3) + (RowNumber + 1) + ",\"NO BUDGET\")", SheetNumber, "#,##0 %");
                        sSheet.Set_Formula(RowNumber, 5, "=VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!$A:D,4,FALSE)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 6, "=VLOOKUP(" + sSheet.GetExcelColumnName(1) + (RowNumber + 1) + ",'CURRENT MONTH TURNOVER SUMMARY'!$A:F,6,FALSE)", SheetNumber, "#,##0");

                        sSheet.Set_Cell(RowNumber, 8, Classes.Global.ConvertToString(newYTDTable.Rows[i]["Name"]).Trim(), SheetNumber);
                        sSheet.Set_Formula(RowNumber, 9, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(9) + (RowNumber + 1) + ",'YTD SALES'!A:C,3,FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 10, "=IFERROR(VLOOKUP(" + sSheet.GetExcelColumnName(9) + (RowNumber + 1) + ",'" + Year + " BUDGET'!A:O,14,FALSE), 0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 11, "=" + sSheet.GetExcelColumnName(10) + (RowNumber + 1) + "-" + sSheet.GetExcelColumnName(11) + (RowNumber + 1), SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 12, "=IFERROR(" + sSheet.GetExcelColumnName(12) + (RowNumber + 1) + "/" + sSheet.GetExcelColumnName(11) + (RowNumber + 1) + ",\"NO BUDGET\")", SheetNumber, "#,##0 %");
                        sSheet.Set_Formula(RowNumber, 13, "=VLOOKUP(" + sSheet.GetExcelColumnName(9) + (RowNumber + 1) + ",'YTD SALES'!$A:D,4,FALSE)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, 14, "=VLOOKUP(" + sSheet.GetExcelColumnName(9) + (RowNumber + 1) + ",'YTD SALES'!$A:F,6,FALSE)", SheetNumber, "#,##0");
                        RowNumber++;
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
                    sSheet.Set_Formula(RowNumber, 1, "='CURRENT MONTH TURNOVER SUMMARY'!C" + (MonthlyTotalRowNumber) + "-'DATA FOR PRESENTATION'!B" + (RowNumber - 1), SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 2, "='" + Year + " BUDGET'!" + sSheet.GetExcelColumnName(MonthColumnIndex) + BudgetTotalRow + " - C" + (RowNumber - 1), SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 3, "=B" + (RowNumber + 1) + "-C" + (RowNumber + 1), SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 4, "=D" + (RowNumber + 1) + "/C" + (RowNumber + 1), SheetNumber, "#,##0 %");
                    sSheet.Set_Formula(RowNumber, 5, "='CURRENT MONTH TURNOVER SUMMARY'!D" + (MonthlyTotalRowNumber) + "-F" + (RowNumber - 1), SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 6, "=F" + (RowNumber + 1) + "/B" + (RowNumber + 1), SheetNumber, "#,##0 %");
                    sSheet.Set_Font_Size("A" + (RowNumber + 1), 12, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center, SheetNumber);
                    sSheet.Set_FontColour("A" + (RowNumber + 1) + ":G" + (RowNumber + 1), Colour, Color.White, SheetNumber);

                    sSheet.Set_Cell(RowNumber, 8, "TOTAL OF OTHERS", SheetNumber, DevExpress.Spreadsheet.SpreadsheetHorizontalAlignment.Center);
                    sSheet.Set_Formula(RowNumber, 9, "='YTD SALES'!C" + (YTDTotalRowNumber) + "-'DATA FOR PRESENTATION'!J" + (RowNumber - 1), SheetNumber, "£ #,##0");
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

                    sSheet.Set_Cell_Alignment("B" + (RowNumber - 19) + ":G" + (RowNumber + 1), "DATA FOR PRESENTATION", SpreadsheetHorizontalAlignment.Center, SpreadsheetVerticalAlignment.Center, true);
                    sSheet.Set_Cell_Alignment("J" + (RowNumber - 19) + ":O" + (RowNumber + 1), "DATA FOR PRESENTATION", SpreadsheetHorizontalAlignment.Center, SpreadsheetVerticalAlignment.Center, true);

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
                    sSheet.Set_Column_Width("A", 33.57, "DATA FOR PRESENTATION");
                    sSheet.Set_Column_Width("I", 33.57, "DATA FOR PRESENTATION");

                    /**************************************************************************************************************************
                    * SALES V BUDGET
                    *************************************************************************************************************************/

                    SheetNumber = 6;
                    RowNumber = 0;
                    sSheet.Insert_Worksheet("SALES V BUDGET", SheetNumber);

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
                    * THIS YEAR NEW BUSINESS
                    *************************************************************************************************************************/

                    SheetNumber++;

                    sSheet.Add_Worksheet(Year + " NEW BUSINESS");

                    RowNumber = 0;

                    sSheet.Set_Cell(RowNumber, 1, "TO BE MANUALLY FILLED IN", SheetNumber);

                    sSheet.Auto_fit(0, 20, SheetNumber);

                    /**************************************************************************************************************************
                    * MONTH SALES PER CUSTOMER SHEETS
                    **************************************************************************************************************************/

                    SheetNumber++;
                    EndDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddMonths(1).AddSeconds(-1).ToString("yyyy-MM-dd");
                    MonthSalesPerCustomerSheets(sSheet, SheetNumber, Year, EndDate, LightGreen);

                    int ThisYearRowCount = sSheet.GetWorksheetRange(Year + " MONTH SALES PER CUSTOMER").RowCount;

                    SheetNumber++;
                    EndDate = LastYear + "-12-31";
                    MonthSalesPerCustomerSheets(sSheet, SheetNumber, LastYear, EndDate, LightGreen);

                    int LastYearRowCount = sSheet.GetWorksheetRange(LastYear + " MONTH SALES PER CUSTOMER").RowCount;

                    SheetNumber++;
                    LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-2).ToString("yyyy");
                    EndDate = LastYear + "-12-31";
                    MonthSalesPerCustomerSheets(sSheet, SheetNumber, LastYear, EndDate, LightGreen);

                    int Last2YearRowCount = sSheet.GetWorksheetRange(LastYear + " MONTH SALES PER CUSTOMER").RowCount;

                    // Resetting afterwards
                    LastYear = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-1).ToString("yyyy");
                    EndDate = Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddMonths(1).AddSeconds(-1).ToString("yyyy-MM-dd");

                    /**************************************************************************************************************************
                    * NEW B V BUDGET
                    *************************************************************************************************************************/
                    Color DarkTeal = System.Drawing.ColorTranslator.FromHtml("#D9E1F2");
                    Color GreenAccent = System.Drawing.ColorTranslator.FromHtml("#E2EFDA");

                    SheetNumber = 8;

                    sSheet.Insert_Worksheet("NEW B V BUDGET", SheetNumber);

                    RowNumber = 0;

                    sSheet.Set_Cell(RowNumber, 0, "NEW BUSINESS WON IN " + LastYear + " IMPACTING " + Year, SheetNumber, SpreadsheetHorizontalAlignment.Left);
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

                    for (int i = NewBusinessWonStart; i < NewBusinessWonEnd; i++)
                    {
                        sSheet.Set_Formula(RowNumber, 0, "='" + Year + " BUDGET'!A" + (i + 1), SheetNumber);
                        sSheet.Set_Formula(RowNumber, 1, "='" + Year + " BUDGET'!B" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 2, "='" + Year + " BUDGET'!C" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 3, "='" + Year + " BUDGET'!D" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 4, "='" + Year + " BUDGET'!E" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 5, "='" + Year + " BUDGET'!F" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 6, "='" + Year + " BUDGET'!G" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 7, "='" + Year + " BUDGET'!H" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 8, "='" + Year + " BUDGET'!I" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 9, "='" + Year + " BUDGET'!J" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 10, "='" + Year + " BUDGET'!K" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 11, "='" + Year + " BUDGET'!L" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 12, "='" + Year + " BUDGET'!M" + (i + 1), SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 13, "=SUM(B" + (RowNumber + 1) + ":M" + (RowNumber + 1) + ")", SheetNumber, "£ #,##0.00");
                        sSheet.Set_Formula(RowNumber, 14, "=SUM(B" + (RowNumber + 1) + ":" + sSheet.GetExcelColumnName(MonthColumnIndex) + (RowNumber + 1) + ")", SheetNumber, "£ #,##0.00");

                        RowNumber++;
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

                    int MonthlyBudgetRow = RowNumber;

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

                    for (int i = newBusinessStart; i < (BudgetTotalRow - 2); i++)
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
                        sSheet.Set_Formula(RowNumber, 12, "='" + Year + " BUDGET'!N" + (i + 1), SheetNumber, "£ #,##0");
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
                    sSheet.Set_Formula(RowNumber, 10, "=IFERROR(SUM(K" + (newBusinessWonSales + 1) + "K" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
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

                    for (int i = newBusinessStart; i < (BudgetTotalRow - 2); i++)
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

                    sSheet.Set_Cell(RowNumber, 0, "MONTHLY SALES - 'NEW IN 2024' WITH BUDGET", SheetNumber, SpreadsheetHorizontalAlignment.Right);
                    sSheet.Set_Formula(RowNumber, 1, "=IFERROR(SUM(B" + (newBusinessSales + 1) + ":B" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 2, "=IFERROR(SUM(C" + (newBusinessSales + 1) + ":C" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 3, "=IFERROR(SUM(D" + (newBusinessSales + 1) + ":D" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 4, "=IFERROR(SUM(E" + (newBusinessSales + 1) + ":E" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 5, "=IFERROR(SUM(F" + (newBusinessSales + 1) + ":F" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 6, "=IFERROR(SUM(G" + (newBusinessSales + 1) + ":G" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 7, "=IFERROR(SUM(H" + (newBusinessSales + 1) + ":H" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 8, "=IFERROR(SUM(I" + (newBusinessSales + 1) + ":I" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 9, "=IFERROR(SUM(J" + (newBusinessSales + 1) + ":J" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
                    sSheet.Set_Formula(RowNumber, 10, "=IFERROR(SUM(K" + (newBusinessSales + 1) + "K" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
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

                    for (int i = 0; i < (newCustomerNoBudgetTotalRows - 1); i++)
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
                    sSheet.Set_Formula(RowNumber, 10, "=IFERROR(SUM(K" + (newBusinessnoBudget + 1) + "K" + (RowNumber) + "),0)", SheetNumber, "£ #,##0");
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

                    // New Business Won
                    sms.Set_Chart(sSheet, "B" + (MonthlyBudgetRow + 1) + ":M" + (MonthlyBudgetRow + 1), "B" + (MonthlySalesRow + 1) + ":M" + (MonthlySalesRow + 1), "A" + (MonthlyBudgetRow + 1), "A" + (MonthlySalesRow + 1), 
                        "B1:M1", "B1:M1", "Q2", "W14", SheetNumber, ChartType.ColumnClustered, Color.DarkGray, Color.Gray, LightGreen, LightGreen, LegendPosition.Bottom, "NEW BUSINESS WON IN " + LastYear + " IMPACTING " + Year);

                    // Budgeted New Business
                    sms.Set_Chart(sSheet, "B" + (newBusinessBudgetEnd + 1) + ":M" + (newBusinessBudgetEnd + 1), "B" + (MonthlySaleswBudget + 1) + ":M" + (MonthlySaleswBudget + 1), "A" + (newBusinessBudgetEnd + 1), "A" + (MonthlySaleswBudget + 1),
                        "B1:M1", "B1:M1", "Q16", "W26", SheetNumber, ChartType.ColumnClustered, Color.DarkGray, Color.Gray, LightGreen, LightGreen, LegendPosition.Bottom, "BUDGETED NEW BUSINESS IN " + Year);

                    // Total New Business
                    sms.Set_Chart(sSheet, "B" + (MonthlySaleswBudget + 1) + ":M" + (MonthlySaleswBudget + 1), "B" + (totalNewNoBudgetRow + 1) + ":M" + (totalNewNoBudgetRow + 1), "A" + (MonthlySaleswBudget + 1), "A" + (totalNewNoBudgetRow + 1),
                        "B1:M1", "B1:M1", "Q28", "Z44", SheetNumber, ChartType.ColumnClustered, Color.Gray, Color.Teal, LightGreen, Color.DarkGray, LegendPosition.Bottom, "TOTAL NEW BUSINESS IN " + Year + " v BUDGET", 
                        "B" + (totalVBudgetRow + 1) + ":M" + (totalVBudgetRow + 1), "A" + (totalVBudgetRow + 1), "B1:M1", "B" + (newBusinessBudgetEnd + 1) + ":M" + (newBusinessBudgetEnd + 1), "A" + (newBusinessBudgetEnd + 1), "B1:M1", true);

                    sSheet.Set_Column_Width(0, 41.57, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(1, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(2, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(3, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(4, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(5, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(6, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(7, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(8, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(9, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(10, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(11, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(12, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(13, 10.86, "NEW B V BUDGET");
                    sSheet.Set_Column_Width(14, 10.86, "NEW B V BUDGET");

                    /**************************************************************************************************************************
                    * SALES V PRIOR YEARS
                    *************************************************************************************************************************/

                    RowNumber = 0;
                    SheetNumber++;

                    sSheet.Insert_Worksheet("SALES V PRIOR YEARS", SheetNumber);

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
                        "B1:M1", "B1:M1", "C25", "J40", SheetNumber, ChartType.ColumnClustered, LightGreen, Color.Teal, Color.Gray, Color.Black, LegendPosition.Bottom, 
                        Classes.Global.ConvertToDateTime(dteReportDate.EditValue).AddYears(-2).ToString("yyyy") + " v " + LastYear + " v " + Year, "B" + (SalesAllYearsRow + 3) + ":M" + (SalesAllYearsRow + 3), 
                        "A" + (SalesAllYearsRow + 3), "B1:M1");

                    //Last Year
                    sms.Set_Chart(sSheet, "B" + (SalesLastYearsRow + 1) + ":M" + (SalesLastYearsRow + 1), "B" + (SalesLastYearsRow + 3) + ":M" + (SalesLastYearsRow + 3), "A" + (SalesLastYearsRow + 1), "A" + (SalesLastYearsRow + 3), 
                        "B1:M1", "B1:M1", "O1", "AA16", SheetNumber, ChartType.ColumnClustered, Color.Teal, Color.Gray, LightGreen, Color.Black, LegendPosition.Bottom, LastYear + " SALES V BUDGET V PRIOR YEAR", 
                        "B" + (SalesLastYearsRow + 2) + ":M" + (SalesLastYearsRow + 2), "A" + (SalesLastYearsRow + 2), "B1:M1", null, null, null, true);

                    //Year
                    sms.Set_Chart(sSheet, "B" + (SalesThisYearRow + 1) + ":M" + (SalesThisYearRow + 1), "B" + (SalesThisYearRow + 3) + ":M" + (SalesThisYearRow + 3), "A" + (SalesThisYearRow + 1), "A" + (SalesThisYearRow + 3), 
                        "B1:M1", "B1:M1", "O18", "AA37", SheetNumber, ChartType.ColumnClustered, Color.Teal, Color.Gray, LightGreen, Color.Black, LegendPosition.Bottom, Year + " SALES V BUDGET V PRIOR YEAR", 
                        "B" + (SalesThisYearRow + 2) + ":M" + (SalesThisYearRow + 2), "A" + (SalesThisYearRow + 2), "B1:M1", null, null, null, true);

                    sSheet.Set_Bold_Range("A1:A50", true, SheetNumber);
                    sSheet.Set_Column_Width(0, 18.29, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(1, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(2, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(3, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(4, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(5, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(6, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(7, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(8, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(9, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(10, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(11, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(12, 12, "SALES V PRIOR YEARS");
                    sSheet.Set_Column_Width(13, 12, "SALES V PRIOR YEARS");

                    /**************************************************************************************************************************
                    * THIS YR V LAST YR
                    *************************************************************************************************************************/

                    Color Pink = ColorTranslator.FromHtml("#FFCCFF");
                    SheetNumber++;
                    RowNumber = 0;

                    sSheet.Insert_Worksheet("THIS YR V LAST YR", SheetNumber);

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
                        if (Classes.Global.ConvertToString(row["Name"]) != "")
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
                        sSheet.Set_Formula(RowNumber, (i + 3), "=IFERROR((" + sSheet.GetExcelColumnName(i + 2) + RowNumber + "-" + sSheet.GetExcelColumnName(i + 3) + (RowNumber) + ")/" + sSheet.GetExcelColumnName(i + 3) + (RowNumber) + ",0)", SheetNumber, "£ #,##0");
                        sSheet.Set_Formula(RowNumber, (i + 4), "=IFERROR((" + sSheet.GetExcelColumnName(i + 2) + RowNumber + "-" + sSheet.GetExcelColumnName(i + 1) + (RowNumber) + ")/" + sSheet.GetExcelColumnName(i + 1) + (RowNumber) + ",0)", SheetNumber, "£ #,##0");

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

                    sSheet.Set_Column_Width(0, 59.00, "THIS YR V LAST YR");

                    for (int i = 1; i < 67; i++)
                        sSheet.Set_Column_Width(i, 13.57, "THIS YR V LAST YR");

                    sSheet.Hide_Columns(StartIndex, 60, SheetNumber);

                    /**************************************************************************************************************************
                    * SALES + WEIGHT EXPORT
                    *************************************************************************************************************************/
                    RowNumber = 0;
                    SheetNumber = 14;
                    sSheet.Add_Worksheet("SALES + WEIGHT EXPORT");

                    sSheet.Set_Cell(RowNumber, 1, "VALUE", SheetNumber);
                    sSheet.Set_Cell(RowNumber, 2, "WEIGHT", SheetNumber);

                    RowNumber++;

                    foreach (DataRow row in YTDSales.Rows)
                    {
                        if (Classes.Global.ConvertToString(row["Name"]) != "")
                        {
                            sSheet.Set_Cell(RowNumber, 0, Classes.Global.ConvertToString(row["Name"]), SheetNumber);
                            sSheet.Set_Formula(RowNumber, 1, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + (((MonthColumnIndex - 1) * 3) - 1) + ",FALSE),0)", SheetNumber, "£ #,##0");
                            sSheet.Set_Formula(RowNumber, 2, "=IFERROR(VLOOKUP(A" + (RowNumber + 1) + ",'" + Year + " MONTH SALES PER CUSTOMER'!A:AZ," + ((MonthColumnIndex - 1) * 3) + ",FALSE),0)", SheetNumber, "#,##0");

                            RowNumber++;
                        }
                    }

                    sSheet.Auto_fit(0, 2, SheetNumber);
                }
                catch(Exception ex)
                {
                    ProcessError.Show(ModuleName, "cmdGo_Click", ex);
                }
                finally
                {
                    /**************************************************************************************************************************
                    * FINAL BITS AND SAVING
                    *************************************************************************************************************************/

                    sSheet.Set_Active_Sheet(0);

                    sSheet.Hide_Wait();

                    // Save the sheet 
                    using (SaveFileDialog saveDialog = new SaveFileDialog())
                    {
                        saveDialog.Filter = "Excel (2010) (.xlsx)|*.xlsx";
                        saveDialog.Title = "Save Spreadsheet";
                        saveDialog.FileName = "Sams Management Sheet " + System.DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss");
                        if (saveDialog.ShowDialog() != DialogResult.Cancel)
                        {
                            sSheet.SaveToFile(saveDialog.FileName, DXTools.Spreadsheet.FormatTypes.Xlsx);
                            System.Diagnostics.Process.Start(saveDialog.FileName);
                        }
                    }
                }
            }
        }

        private void MonthSalesPerCustomerSheets(Spreadsheet sSheet, int SheetNumber, string Year, string EndDate, Color LightGreen)
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
                        "SUM(tbl_Product.Unit_Weight * tbl_InvoiceItem.Qty_Order) AS Line_Unit_Weight, " +
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

                sSheet.Set_Cell(RowNumber, 0, Year, SheetNumber, SpreadsheetHorizontalAlignment.Center);
                sSheet.Set_Rotation("A1", SheetNumber, 0, SpreadsheetVerticalAlignment.Center);
                sSheet.Set_Font_Size("A1", 20, SpreadsheetHorizontalAlignment.Center, SheetNumber);
                sSheet.Set_Bold_Range("A1:AN1", true, SheetNumber);
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

                RowNumber++;

                sSheet.FormatCell("B:AN", "£ #,##0", SheetNumber);

                foreach (var priorYearSalesGroup in PriorYearSaleList.GroupBy(g => new { g.Name, g.Deleted }))
                {
                    sSheet.Set_Cell(RowNumber, 0, priorYearSalesGroup.Key.Name, SheetNumber);
                    if (priorYearSalesGroup.Key.Deleted)
                        sSheet.Set_FontColour("A" + (RowNumber + 1) + ":AK" + (RowNumber + 1), Color.Gray, Color.White, SheetNumber);
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
                    RowNumber++;
                }

                sSheet.Set_Column_Width(0, 46.43, Year + " MONTH SALES PER CUSTOMER");

                for (int i = 1; i <= 39; i++)
                    sSheet.Set_Column_Width(i, 11.86, Year + " MONTH SALES PER CUSTOMER");

                sSheet.Set_FontColour("AL1:AN" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);
                sSheet.Set_AllBorders("A1:AN" + RowNumber, Color.Black, BorderLineStyle.Thin, SheetNumber);
                sSheet.Set_OutsideBorders("A1:A" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
                sSheet.Set_OutsideBorders("E1:G" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
                sSheet.Set_OutsideBorders("K1:M" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
                sSheet.Set_OutsideBorders("Q1:S" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
                sSheet.Set_OutsideBorders("W1:Y" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
                sSheet.Set_OutsideBorders("AC1:AE" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
                sSheet.Set_OutsideBorders("AI1:AK" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);
                sSheet.Set_OutsideBorders("A1:AN" + RowNumber, Color.Black, SheetNumber, BorderLineStyle.Medium);

                sSheet.Set_Cell(RowNumber, 0, "TOTAL", SheetNumber, SpreadsheetHorizontalAlignment.Left);

                for (int i = 1; i < 39; i += 3)
                {
                    sSheet.Set_Formula(RowNumber, i, "=SUM(" + sSheet.GetExcelColumnName(i + 1) + "3:" + sSheet.GetExcelColumnName(i + 1) + (RowNumber) + ")", SheetNumber, "£ #,##0.00");
                    sSheet.Set_Formula(RowNumber, (i + 1), "=SUM(" + sSheet.GetExcelColumnName(i + 2) + "3:" + sSheet.GetExcelColumnName(i + 2) + (RowNumber) + ")", SheetNumber);
                    sSheet.Set_Formula(RowNumber, (i + 2), "=IFERROR(" + sSheet.GetExcelColumnName(i + 1) + (RowNumber + 1) + "/" + sSheet.GetExcelColumnName(i + 2) + (RowNumber + 1) + ",0)", SheetNumber, "£ #,##0.00");
                }

                sSheet.Set_AllBorders("A" + (RowNumber + 1) + ":AN" + (RowNumber + 1), Color.Black, SheetNumber);
                sSheet.Set_OutsideBorders("A" + (RowNumber + 1) + ":AN" + (RowNumber + 1), Color.Black, SheetNumber, BorderLineStyle.Medium);
                sSheet.Set_FontColour("A" + (RowNumber + 1) + ":AN" + (RowNumber + 1), Color.LightGray, Color.Black, SheetNumber);

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
            }
            catch (Exception ex)
            {
                ProcessError.Show(ModuleName, "MonthSalesPerCustomerSheets", ex);
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
                ProcessError.Show(ModuleName, "ConvertBudgetSheetToDataTable", ex);
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
    }
}