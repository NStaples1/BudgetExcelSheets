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
            Color LightGreen = ColorTranslator.FromHtml("#66FFCC");

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

                List<Color> colors = new List<Color> { LightGreen, Color.Teal, Color.Gray };

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

        public void Set_Chart(Spreadsheet sSheet, string FirstCellReferenceRange, string SecondCellReferenceRange, string FirstTitle, string SecondTitle, string FirstPlotRange, string SecondPlotRange, string TopLeft, 
            string BottomRight, int SheetIndex, ChartType chartType, Color color1, Color color2, Color color3, Color color4, LegendPosition legendPosition = LegendPosition.Bottom, string Title = "", 
            string ThirdCellReferenceRange = "", string ThirdTitle = "", string ThirdPlotRange = "", string FourthCellReferenceRange = "", string FourthTitle = "", string FourthPlotRange = "", bool Combo = false, bool Legend = true, 
            bool DataTable = false)
        {
            Worksheet workSheet = sSheet.workbook.Worksheets[SheetIndex];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet Index " + SheetIndex);
            else
            {
                Color LightGreen = ColorTranslator.FromHtml("#66FFCC");

                Chart chart = workSheet.Charts.Add(chartType);
                chart.TopLeftCell = workSheet.Cells[TopLeft];
                chart.BottomRightCell = workSheet.Cells[BottomRight];
                chart.Title.Visible = true;
                chart.Title.SetValue(Title);
                chart.Title.Font.Size = 14;
                chart.Views[0].DataLabels.ShowValue = false;
                chart.Views[0].VaryColors = false;
                chart.Legend.Visible = Legend;
                chart.Legend.Position = legendPosition;
                chart.DataTable.Visible = DataTable;
                
                List<Color> colors = new List<Color> { LightGreen, Color.Teal, Color.Gray };
                Series series0;
                Series series1;
                Series series2;
                Series series3;

                switch (chartType)
                {
                    default:
                        if (!Combo)
                        {
                            // Add the data to the chart
                            // Asking for the Series Name, the plot data e.g Jan, Feb etc. and the Values needed
                            series0 = chart.Series.Add(workSheet[FirstTitle], workSheet[FirstPlotRange], workSheet[FirstCellReferenceRange]);
                            series1 = chart.Series.Add(workSheet[SecondTitle], workSheet[SecondPlotRange], workSheet[SecondCellReferenceRange]);
                            if (ThirdCellReferenceRange != "")
                            {
                                if (ThirdCellReferenceRange != null)
                                {
                                    series2 = chart.Series.Add(workSheet[ThirdTitle], workSheet[ThirdPlotRange], workSheet[ThirdCellReferenceRange]);
                                    chart.Series[2].Fill.SetSolidFill(color3);
                                }
                            }

                            if (FourthCellReferenceRange != "")
                            {
                                if (FourthCellReferenceRange != null)
                                {
                                    series3 = chart.Series.Add(workSheet[FourthTitle], workSheet[FourthPlotRange], workSheet[FourthCellReferenceRange]);
                                    chart.Series[3].Fill.SetSolidFill(color4);
                                }
                            }

                            chart.Series[0].Fill.SetSolidFill(color1);
                            chart.Series[1].Fill.SetSolidFill(color2);
                        }
                        else
                        {
                            series0 = chart.Series.Add(workSheet[FirstTitle], workSheet[FirstPlotRange], workSheet[FirstCellReferenceRange]);
                            series1 = chart.Series.Add(workSheet[SecondTitle], workSheet[SecondPlotRange], workSheet[SecondCellReferenceRange]);
                            series2 = chart.Series.Add(workSheet[ThirdTitle], workSheet[ThirdPlotRange], workSheet[ThirdCellReferenceRange]);
                            if (FourthCellReferenceRange != "")
                            {
                                if (FourthCellReferenceRange != null)
                                {
                                    series3 = chart.Series.Add(workSheet[FourthTitle], workSheet[FourthPlotRange], workSheet[FourthCellReferenceRange]);
                                    chart.Series[3].ChangeType(ChartType.Line);
                                    chart.Series[3].Outline.SetSolidFill(color4);
                                }
                                else
                                {
                                    chart.Series[2].ChangeType(ChartType.Line);
                                    chart.Series[2].Outline.SetSolidFill(LightGreen);
                                }
                            }
                            else
                            {
                                chart.Series[2].ChangeType(ChartType.Line);
                                chart.Series[2].Outline.SetSolidFill(LightGreen);
                            }

                            chart.Series[0].Fill.SetSolidFill(color1);
                            chart.Series[1].Fill.SetSolidFill(color2);
                            // If its a line it wont get used
                            chart.Series[2].Fill.SetSolidFill(color3);

                        }
                        break;
                }
            }
        }


        #endregion
    }
}


