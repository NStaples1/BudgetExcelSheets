using System;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using DXTools.Classes;
using System.Windows.Forms;
using DXTools.Models;
using DevExpress.XtraSplashScreen;
using System.Drawing;
using System.Reflection;
using System.Drawing.Imaging;
using DevExpress.Spreadsheet.Charts;
using System.Xml;
using DevExpress.XtraSpreadsheet.Commands;
using DevExpress.Spreadsheet.Drawings;
using DevExpress.XtraSpreadsheet.DocumentFormats.Xlsb;
using DXTools;

namespace DXTools.Classes
{
    public class SamsManagementSheet
    {
        private const string ModuleName = "DXTools.Classes.SamsManagementSheet";
        public enum FormatTypes
        {
            Csv,
            Xls,
            Xlsx,
            PDF
        }

        public enum ObjectTypes
        {
            Double,
            Int,
            String,
            DateTime,
            Boolean
        }

        #region Charts

        public void Set_Pie_Chart(Spreadsheet sSheet, string CellReferenceRange, ChartType chartType, int SheetIndex, string TopLeft, string BottomRight, string Title = "", LegendPosition legendPosition = LegendPosition.Bottom)
        {
            Worksheet workSheet = sSheet.workbook.Worksheets[SheetIndex];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet Index " + SheetIndex);
            else
            {
                Chart chart = workSheet.Charts.Add(chartType);
                chart.SelectData(workSheet[CellReferenceRange], ChartDataDirection.Row);
                chart.TopLeftCell = workSheet.Cells[TopLeft];
                chart.BottomRightCell = workSheet.Cells[BottomRight];
                chart.Title.Visible = true;
                chart.Title.SetValue(Title);
                chart.Title.Font.Size = 10;
                List<Color> colors = new List<Color> { Color.Teal, Color.SeaGreen, Color.LightGray };

                chart.Views[0].DataLabels.ShowValue = true;
                chart.Views[0].DataLabels.Font.Bold = true;
                chart.Views[0].VaryColors = true;
                chart.Legend.Visible = true;
                chart.Legend.Position = legendPosition;
                foreach (var series in chart.Series)
                {
                    series.Explosion = 3;
                    series.CustomDataPoints.Add(0).Fill.SetSolidFill(colors[0]);
                    series.CustomDataPoints.Add(1).Fill.SetSolidFill(colors[1]);
                    series.CustomDataPoints.Add(2).Fill.SetSolidFill(colors[2]);
                }
            }
        }

        public void Set_Chart(Spreadsheet sSheet, string CellReferenceRange, ChartType chartType, int SheetIndex, string TopLeft, string BottomRight, string Title = "", LegendPosition legendPosition = LegendPosition.Bottom)
        {
            Worksheet workSheet = sSheet.workbook.Worksheets[SheetIndex];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet Index " + SheetIndex);
            else
            {
                Chart chart = workSheet.Charts.Add(chartType);
                chart.SelectData(workSheet[CellReferenceRange], ChartDataDirection.Row);
                chart.TopLeftCell = workSheet.Cells[TopLeft];
                chart.BottomRightCell = workSheet.Cells[BottomRight];
                chart.Title.Visible = true;
                chart.Title.SetValue(Title);
                chart.Title.Font.Size = 10;
                chart.Views[0].DataLabels.ShowValue = false;
                chart.Views[0].VaryColors = true;
                chart.Legend.Visible = true;
                chart.Legend.Position = legendPosition;
                Series series1 = chart.Series.Add(
                new CellValue[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" },
                new CellValue[] { 50, 100, 30, 104, 87, 150 });
            }
        }


        #endregion
    }
}


