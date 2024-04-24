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
using DevExpress.UnitConversion;
using DevExpress.XtraSpreadsheet.DocumentFormats.Xlsb;

namespace DXTools
{
    public class Spreadsheet : IDisposable
    {
        internal Workbook workbook = new Workbook();
        private const string ModuleName = "DXTools.Classes.Spreadsheet";
        public List<ExceptionLog> ExceptionLogs = new List<ExceptionLog>();

        public void Dispose()
        {
            if (workbook != null)
            {
                workbook.Dispose();
                workbook = null;
            }
        }

        public Spreadsheet()
        {
            ThrowExceptionOnError = true;
        }

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

        public bool ThrowExceptionOnError { get; set; }

        #region Load / Import

        public bool LoadFromFile(string Filename)
        {
            try
            {
                workbook = null;
                workbook = new Workbook();
                workbook.LoadDocument(Filename);

                //// TODO: some error logging
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ProcessError.Return_Error(ModuleName, "Load", ex));
            }
        }

        public void Import(DataTable SourceData, bool AddHeader = true, int FirstRowIndex = 0, int FirstColumnIndex = 0, int SheetNumber = 0)
        {
            try
            {
                Worksheet worksheet = workbook.Worksheets[SheetNumber];

                DataImportOptions Options = new DataImportOptions();
                worksheet.Import(SourceData, AddHeader, FirstRowIndex, FirstColumnIndex, Options);
            }
            catch (Exception ex)
            {
                throw new Exception(ProcessError.Return_Error(ModuleName, "Import", ex));
            }
        }

        public void SaveToFile(string Filename, FormatTypes Format)
        {
            try
            {
                using (FileStream stream = new FileStream(Filename, FileMode.Create, FileAccess.ReadWrite))
                {
                    switch (Format)
                    {
                        case FormatTypes.Csv:
                            workbook.SaveDocument(stream, DocumentFormat.Csv);
                            break;
                        case FormatTypes.Xls:
                            workbook.SaveDocument(stream, DocumentFormat.Xls);
                            break;
                        case FormatTypes.Xlsx:
                            workbook.SaveDocument(stream, DocumentFormat.Xlsx);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ProcessError.Return_Error(ModuleName, "SaveToFile", ex));
            }
        }

        #endregion
        #region Exports

        public DataTable Export_To_Datatable()
        {
            return Export_To_Datatable(workbook.Worksheets.ActiveWorksheet.Name);
        }

        public DataTable Export_To_Datatable(int SheetIndex)
        {
            Worksheet workSheet = workbook.Worksheets[SheetIndex];
            return Export_To_Datatable(workSheet.Name, workSheet.GetUsedRange());
        }

        public DataTable Export_To_Datatable(int SheetIndex, bool HasHeaders)
        {
            Worksheet workSheet = workbook.Worksheets[SheetIndex];
            return Export_To_Datatable(workSheet.Name, workSheet.GetUsedRange(), HasHeaders);
        }

        public DataTable Export_To_Datatable(int SheetIndex, CellRange mRange)
        {
            Worksheet workSheet = workbook.Worksheets[SheetIndex];
            return Export_To_Datatable(workSheet.Name, mRange);
        }

        public DataTable Export_To_Datatable(string SheetName, CellRange mRange = null, bool HasHeaders = true)
        {
            Worksheet workSheet = workbook.Worksheets[SheetName];

            if (mRange == null)
                mRange = workSheet.GetUsedRange();

            DataTable dataTable = workSheet.CreateDataTable(mRange, true);
            DataTableExporter exporter = workSheet.CreateDataTableExporter(mRange, dataTable, HasHeaders);
            // Specify exporter options.
            exporter.Options.ConvertEmptyCells = true;
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = true;
            exporter.CellValueConversionError += exporter_CellValueConversionError;

            // Perform the export.
            exporter.Export();

            return dataTable;
        }

        private void exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e)
        {
            e.DataTableValue = null;

            if (ThrowExceptionOnError)
            {
                e.Action = DataTableExporterAction.Stop;
                throw new Exception("Error in cell " + e.Cell.GetReferenceA1());
            }
            else
            {
                ExceptionLog eLog = new ExceptionLog() { CellReference = e.Cell.GetReferenceA1(), ExceptionValue = e.CellValue.ToString() };
                ExceptionLogs.Add(eLog);
                e.Action = DataTableExporterAction.SkipRow;
            }
        }

        public Stream Export_To_Stream(FormatTypes formatType)
        {
            try
            {
                MemoryStream memStream = new MemoryStream();

                switch (formatType)
                {
                    case FormatTypes.Xls:
                        if (workbook != null)
                        {
                            workbook.CalculateFull();
                            workbook.SaveDocument(memStream, DocumentFormat.Xls);
                        }
                        break;
                    case FormatTypes.Xlsx:
                        if (workbook != null)
                        {
                            workbook.CalculateFull();
                            workbook.SaveDocument(memStream, DocumentFormat.Xlsx);
                        }
                        break;
                    case FormatTypes.PDF:
                        if (workbook != null)
                        {
                            workbook.CalculateFull();
                            workbook.ExportToPdf(memStream);
                        }
                        break;
                }
                memStream.Flush(); //Always catches me out
                memStream.Position = 0;

                return memStream;
            }
            catch (Exception ex)
            {
                // TODO: some error logging
                return null;
            }
        }


        #endregion
        #region Range
        public enum DocumentUnits
        {
            Point,
            Millimetres,
            Inch,
            Document,
            Centimetre
        }

        public void Set_Workbook_Units(DocumentUnits Units)
        {
            switch (Units)
            {
                case DocumentUnits.Point:
                    workbook.Unit = DevExpress.Office.DocumentUnit.Point; 
                    break;
                case DocumentUnits.Millimetres:
                    workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
                    break;
                case DocumentUnits.Centimetre:
                    workbook.Unit = DevExpress.Office.DocumentUnit.Centimeter;
                    break;
                case DocumentUnits.Inch:
                    workbook.Unit = DevExpress.Office.DocumentUnit.Inch;
                    break;
                case DocumentUnits.Document:
                    workbook.Unit = DevExpress.Office.DocumentUnit.Document;
                    break;

            }
        }

      public string GetExcelRange(int Row, int Column)
      {
         /**************************************************************************************
          * The get excel col doesnt work off 0 based indexes so add 1
          *************************************************************************************/
         return GetExcelColumnName(Column + 1) + Row;
      }


      public string GetExcelRange(int StartRow, int StartColumn, int LastRow, int LastColumn)
      {
         /**************************************************************************************
          * The get excel col doesnt work off 0 based indexes so add 1
          *************************************************************************************/
         return GetExcelColumnName(StartColumn + 1) + StartRow + ":" + GetExcelColumnName(LastColumn + 1) + LastRow;
      }

      public CellRange GetExcelCellRange(int StartRow, int StartColumn, int LastRow, int LastColumn, int SheetIndex = 0)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         return workSheet.Range[GetExcelColumnName(StartColumn) + StartRow + ":" + GetExcelColumnName(LastColumn) + LastRow];
      }

      public string GetExcelColumnName(int columnNumber)
      {
         int dividend = columnNumber;
         string columnName = String.Empty;
         int modulo;

         while (dividend > 0)
         {
            modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            dividend = (int)((dividend - modulo) / 26);
         }

         return columnName;
      }

        public CellRange GetWorksheetRange(string sheetName)
        {
            Worksheet workSheet = workbook.Worksheets[sheetName];
            return workSheet.GetDataRange();
        }

      #endregion
      #region Auto Filter

      public void Auto_Filter(int SheetIndex = 0)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.GetDataRange();
            workSheet.AutoFilter.Apply(range);
         }
      }

      public void Auto_Filter(string CellRange, int SheetIndex = 0)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            workSheet.AutoFilter.Apply(workSheet[CellRange]);
         }
      }

      #endregion
      #region Add Worksheet

      public int Add_Worksheet(string SheetName)
      {
         // Sheet names cant exceed 31 chars
         if (SheetName.Length > 31)
            SheetName = SheetName.Substring(0, 28) + "...";

         Worksheet workSheet = workbook.Worksheets.Add(SheetName);
         if (workSheet == null)
            throw new Exception("Unable to create Sheet " + SheetName);

         return workSheet.Index;
      }

      public int Get_Worksheet_Index(string SheetName)
      {
         Worksheet workSheet = workbook.Worksheets[SheetName];
         if (workSheet != null)
            return workSheet.Index;
         else
            return -1;
      }

      public void Insert_Worksheet(string SheetName, int SheetIndex = 0)
      {
         // Sheet names cant exceed 31 chars
         if (SheetName.Length > 31)
            SheetName = SheetName.Substring(0, 28) + "...";

         Worksheet workSheet = workbook.Worksheets.Insert(SheetIndex, SheetName);
         if (workSheet == null)
            throw new Exception("Unable to create Sheet " + SheetName);
      }

      #endregion
      #region Auto fit

      public void Auto_fit(int FirstCol, int LastCol, string SheetName)
      {
         Worksheet workSheet = workbook.Worksheets[SheetName];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Name " + SheetName);
         else
            workSheet.Columns.AutoFit(FirstCol, LastCol);
      }

      public void Auto_fit(int FirstCol, int LastCol, int Sheetindex)
      {
         Worksheet workSheet = workbook.Worksheets[Sheetindex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + Sheetindex);
         else
            workSheet.Columns.AutoFit(FirstCol, LastCol);
      }

      #endregion
      #region FormatCell

      public void FormatCell(string CellReferenceRange, string FormatString, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.Range[CellReferenceRange];
            range.NumberFormat = FormatString;
         }
      }

        public void Set_Column_Width(string ColumnName, double Width, string SheetName)
        {
            Worksheet workSheet = workbook.Worksheets[SheetName];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet Name " + SheetName);
            else
                workSheet.Columns[ColumnName].WidthInCharacters = Width;
        }

        public void Set_Column_Width(int ColumnIndex, double Width, string SheetName)
        {
            Worksheet workSheet = workbook.Worksheets[SheetName];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet Name " + SheetName);
            else
                workSheet.Columns[ColumnIndex].WidthInCharacters = Width;
        }

        public void Set_Row_Height(int RowIndex, double Height, string SheetName)
        {
            Worksheet workSheet = workbook.Worksheets[SheetName];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet Name " + SheetName);
            else
                workSheet.Rows[RowIndex].Height = Height;
        }


        #endregion
        #region Freeze Planes

        public void FreezePlanes(int RowIndex, int ColIndex, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange visibleRange = workSheet.GetUsedRange();
            workSheet.FreezePanes(RowIndex, ColIndex, visibleRange);
         }
      }

      public void FreezeRows(int RowIndex, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            workSheet.FreezeRows(RowIndex);
         }
      }

      #endregion
      #region Set Bold

      public void Set_Bold(string CellReference, bool FontBold, string SheetName)
      {
         Worksheet workSheet = workbook.Worksheets[SheetName];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Name " + SheetName);
         else
            workSheet.Cells[CellReference].Font.Bold = FontBold;
      }

      public void Set_Bold(int RowIndex, int ColumnIndex, bool FontBold, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
            workSheet.Cells[RowIndex, ColumnIndex].Font.Bold = FontBold;
      }

      public void Set_Bold(string CellReference, bool FontBold, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
            workSheet.Cells[CellReference].Font.Bold = FontBold;
      }

      public void Set_Bold(int RowIndex, int ColumnIndex, bool FontBold, string SheetName)
      {
         Worksheet workSheet = workbook.Worksheets[SheetName];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Name " + SheetName);
         else
            workSheet.Cells[RowIndex, ColumnIndex].Font.Bold = FontBold;
      }

      public void Set_Bold_Range(string cellRange, bool IsBold, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet index " + SheetIndex);
         else
         {
            CellRange mRange = workSheet.Range[cellRange];
            DevExpress.Spreadsheet.Formatting newFmt = mRange.BeginUpdateFormatting();
            try
            {
               if (IsBold)
                  newFmt.Font.FontStyle = SpreadsheetFontStyle.Bold;
               else
                  newFmt.Font.FontStyle = SpreadsheetFontStyle.Regular;
            }
            finally
            {
               mRange.EndUpdateFormatting(newFmt);
            }
         }
      }

      #endregion
      #region Get Cell

      public List<CellModel> GetCellFormattingRange(int StartRow, int StartColumn, int LastRow, int LastColumn, int SheetIndex = 0)
      {
         return GetCellFormattingRange(GetExcelCellRange(StartRow, StartColumn, LastRow, LastColumn, SheetIndex));
      }

      public List<CellModel> GetCellFormattingRange(CellRange Range)
      {
         List<CellModel> FormattingList = new List<CellModel>();
         foreach (Cell cell in Range.ExistingCells)
         {
            FormattingList.Add(GetCellFormatting(cell, cell.Worksheet));
         }
         return FormattingList;
      }

      public CellModel GetCellFormatting(int RowIndex, int ColumnIndex, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
            return GetCellFormatting(workSheet.Cells[RowIndex, ColumnIndex], workSheet);
      }

      public CellModel GetCellFormatting(Cell cell, Worksheet workSheet)
      {
         CellModel cm = new CellModel();

         if (cell != null)
         {
            cm.RowIndex = cell.RowIndex;
            cm.ColumnIndex = cell.ColumnIndex;
            cm.SheetIndex = workSheet.Index;

            cm.RowReference = (cm.ColumnIndex + 1).ToString();
            cm.ColumnReference = GetExcelColumnName(cm.ColumnIndex);
            cm.SheetName = workSheet.Name;
            cm.Forecolour = cell.Font.Color;
            cm.Backcolour = cell.FillColor;
         }

         return cm;
      }

      public int GetEndCellReference(int RowIndex, int ColumnIndex, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            Cell cell = workSheet.Cells[RowIndex, ColumnIndex];

            if (cell.IsMerged)
            {
               IList<CellRange> MergedRange = cell.GetMergedRanges();
               if (MergedRange.Count > 0)
                  return MergedRange[0].ExistingCells.LastOrDefault().ColumnIndex;
               else
                  return ColumnIndex;
            }
            else
               return ColumnIndex;
         }
      }

      public CellValue Get_Cell_Value(int RowIndex, int ColumnIndex, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            Cell cell = workSheet.Cells[RowIndex, ColumnIndex];
            return cell.Value;
         }
      }

      public string Get_Cell_Text(int RowIndex, int ColumnIndex, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            Cell cell = workSheet.Cells[RowIndex, ColumnIndex];

            if (cell.Value.IsDateTime)
               return Global.ConvertToDateTime(cell.Value.DateTimeValue).ToString("dd/MM/yyyy");
            else if (cell.Value.IsNumeric)
               return Global.ConvertDoubleToString(cell.Value.NumericValue, cell.NumberFormat);
            else
               return cell.Value.TextValue;
         }
      }

      public CellValue Get_Cell_Formula(int RowIndex, int ColumnIndex, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            Cell cell = workSheet.Cells[RowIndex, ColumnIndex];
            return cell.Formula;
         }
      }

        #endregion
        #region Set Cell


        public void Set_Cell_Alignment(string cellRange, string SheetName, SpreadsheetHorizontalAlignment CellAlignment = SpreadsheetHorizontalAlignment.Left, SpreadsheetVerticalAlignment VertAlign = SpreadsheetVerticalAlignment.Center ,bool WrapText = false)
        {
            Worksheet workSheet = workbook.Worksheets[SheetName];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet Name " + SheetName);
            else
            {
                CellRange mRange = workSheet.Range[cellRange];
                mRange.Alignment.Horizontal = CellAlignment;
                mRange.Alignment.WrapText = WrapText;
            }
        }


        public void Set_Cell(string CellReference, object CellValue, string SheetName, SpreadsheetHorizontalAlignment CellAlignment = SpreadsheetHorizontalAlignment.Left)
      {
         Worksheet workSheet = workbook.Worksheets[SheetName];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Name " + SheetName);
         else
         {
            Cell cell = workSheet.Cells[CellReference];
            cell.Alignment.Horizontal = CellAlignment;
            Set_Cell(cell, CellValue);
         }
      }

      public void Set_Cell(string CellReference, object CellValue, int SheetIndex, SpreadsheetHorizontalAlignment CellAlignment = SpreadsheetHorizontalAlignment.Left)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            Cell cell = workSheet.Cells[CellReference];
            cell.Alignment.Horizontal = CellAlignment;
            Set_Cell(cell, CellValue);
         }
      }

      public void Set_Cell(int RowIndex, int ColumnIndex, object CellValue, string SheetName, SpreadsheetHorizontalAlignment CellAlignment = SpreadsheetHorizontalAlignment.Left)
      {
         Worksheet workSheet = workbook.Worksheets[SheetName];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Name " + SheetName);
         else
         {
            Cell cell = workSheet.Cells[RowIndex, ColumnIndex];
            cell.Alignment.Horizontal = CellAlignment;
            Set_Cell(cell, CellValue);
         }
      }

      public void Set_Cell(int RowIndex, int ColumnIndex, object CellValue, int SheetIndex, SpreadsheetHorizontalAlignment CellAlignment = SpreadsheetHorizontalAlignment.Left, bool wrapText = false)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            Cell cell = workSheet.Cells[RowIndex, ColumnIndex];
            cell.Alignment.Horizontal = CellAlignment;
            cell.Alignment.WrapText = wrapText;
            Set_Cell(cell, CellValue);
         }
      }

      private void Set_Cell(Cell cell, object CellValue)
      {
         if (CellValue == null)
            Set_Cell(cell, CellValue, ObjectTypes.String);
         else
         {
            switch (Type.GetTypeCode(CellValue.GetType()))
            {
               case TypeCode.Int16:
               case TypeCode.Int32:
               case TypeCode.Int64:
                  Set_Cell(cell, CellValue, ObjectTypes.Int);
                  break;
               case TypeCode.Double:
                  Set_Cell(cell, CellValue, ObjectTypes.Double);
                  break;
               case TypeCode.DBNull:
               case TypeCode.String:
                  Set_Cell(cell, CellValue, ObjectTypes.String);
                  break;
               case TypeCode.DateTime:
                  Set_Cell(cell, CellValue, ObjectTypes.DateTime);
                  break;
               case TypeCode.Boolean:
                  Set_Cell(cell, CellValue, ObjectTypes.Boolean);
                  break;
               default:
                  throw new Exception("Unknown Object Type " + CellValue.GetType().ToString());
            }
         }
      }

      private void Set_Cell(Cell cell, object CellValue, ObjectTypes ObjType)
      {
         switch (ObjType)
         {
            case ObjectTypes.Double:
               cell.Value = ObjectConversion.ConvertToDouble(CellValue);
               break;
            case ObjectTypes.Int:
               cell.Value = ObjectConversion.ConvertToInt(CellValue);
               break;
            case ObjectTypes.String:
               cell.Value = ObjectConversion.ConvertToString(CellValue);
               break;
            case ObjectTypes.DateTime:
               cell.Value = ObjectConversion.ConvertToDateTime(CellValue);
               break;
            case ObjectTypes.Boolean:
               cell.Value = ObjectConversion.ConvertToBool(CellValue);
               break;
         }
      }

      public void Set_Formula(int RowIndex, int ColumnIndex, string FormulaValue, int SheetIndex, string formatString = "")
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            Cell cell = workSheet.Cells[RowIndex, ColumnIndex];
            cell.SetValueFromText(FormulaValue);
                if (!string.IsNullOrEmpty(formatString))
                    cell.NumberFormat = formatString;
         }
      }

        #endregion

        

        #region Set Font

        public void Set_Font(string CellReference, FontModel mFont, string SheetName)
      {
         Worksheet workSheet = workbook.Worksheets[SheetName];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Name " + SheetName);
         else
         {
            SpreadsheetFont font = workSheet.Cells[CellReference].Font;
            Set_Font(font, mFont);
         }
      }

      public void Set_Font(int RowIndex, int ColumnIndex, FontModel mFont, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            SpreadsheetFont font = workSheet.Cells[RowIndex, ColumnIndex].Font;
            Set_Font(font, mFont);
         }
      }

      public void Set_Font(SpreadsheetFont cellFont, FontModel mFont)
      {
         cellFont.Name = mFont.FontName;
         cellFont.Size = mFont.FontSize;
         cellFont.Color = mFont.FontColour;
         cellFont.Bold = mFont.Bold;
         if (mFont.Underline)
            cellFont.UnderlineType = UnderlineType.Single;
         else
            cellFont.UnderlineType = UnderlineType.None;
      }

      public void Set_Font_Range(string cellRange, FontModel mFont, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.Range[cellRange];
            range.Font.Name = mFont.FontName;
            range.Font.Size = mFont.FontSize;
            range.Font.Color = mFont.FontColour;
            if (mFont.Underline)
               range.Font.UnderlineType = UnderlineType.Single;
            else
               range.Font.UnderlineType = UnderlineType.None;
         }
      }

      public void Set_Font_Sheet(FontModel mFont, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.GetUsedRange();
            range.Font.Name = mFont.FontName;
            range.Font.Size = mFont.FontSize;
            range.Font.Color = mFont.FontColour;
            if (mFont.Underline)
               range.Font.UnderlineType = UnderlineType.Single;
            else
               range.Font.UnderlineType = UnderlineType.None;
         }
      }


      #endregion
      #region Set Font

      public void Set_Font_Size(string CellReference, int FontSize, SpreadsheetHorizontalAlignment Alignment, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet index " + SheetIndex);
         else
         {
            SpreadsheetFont font = workSheet.Cells[CellReference].Font;
            font.Size = FontSize;

            Cell cell = workSheet.Cells[CellReference];
            cell.Alignment.Horizontal = Alignment;
         }
      }

        public void Set_Rotation(string CellReference, int SheetIndex, int Rotation)
        {
            Worksheet workSheet = workbook.Worksheets[SheetIndex];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet index " + SheetIndex);
            else
            {
                Cell cell = workSheet.Cells[CellReference];
                cell.Alignment.RotationAngle = Rotation;
            }
        }
        public void Set_Rotation(string CellReference, int SheetIndex, int Rotation, SpreadsheetVerticalAlignment Alignment)
        {
            Worksheet workSheet = workbook.Worksheets[SheetIndex];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet index " + SheetIndex);
            else
            {
                Cell cell = workSheet.Cells[CellReference];
                cell.Alignment.RotationAngle = Rotation;
                cell.Alignment.Vertical = Alignment;
            }
        }

        #endregion
        #region Set Font Colour

        public void Set_FontColour(string CellReferenceRange, System.Drawing.Color? BackGroundColour, System.Drawing.Color? ForeGroundColour, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.Range[CellReferenceRange];
            DevExpress.Spreadsheet.Formatting rangeFormatting = range.BeginUpdateFormatting();
            if (ForeGroundColour.HasValue)
               rangeFormatting.Font.Color = ForeGroundColour.Value;
            if (BackGroundColour.HasValue)
               rangeFormatting.Fill.BackgroundColor = BackGroundColour.Value;
            range.EndUpdateFormatting(rangeFormatting);
         }
      }

      #endregion
      #region Set Cell Backcolour

      public string Set_BackColour(string Range, Color Colour, int SheetIndex)
      {
         try
         {
            Worksheet workSheet = workbook.Worksheets[SheetIndex];
            SetRangeBackColor(workSheet[Range], Colour);

            return string.Empty;
         }
         catch (Exception ex)
         {
            return ex.Message;
         }
      }

        public string Set_BackColour(string Range, string Hex, int SheetIndex)
        {
            try
            {
                Worksheet workSheet = workbook.Worksheets[SheetIndex];
                Color Colour = System.Drawing.ColorTranslator.FromHtml(Hex);
                SetRangeBackColor(workSheet[Range], Colour);

                return string.Empty;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string Set_BackColour(int RowIndex, int ColumnIndex, Color Colour, int SheetIndex)
      {
         try
         {
            Worksheet workSheet = workbook.Worksheets[SheetIndex];
            SetRangeBackColor(workSheet[RowIndex, ColumnIndex], Colour);

            return string.Empty;
         }
         catch (Exception ex)
         {
            return ex.Message;
         }
      }

      private void SetRangeBackColor(CellRange currentRange, System.Drawing.Color color)
      {
         DevExpress.Spreadsheet.Formatting newFmt = currentRange.BeginUpdateFormatting();
         try
         {
            newFmt.Fill.BackgroundColor = color;
         }
         finally
         {
            currentRange.EndUpdateFormatting(newFmt);
         }
      }

      #endregion
      #region Conditional Formatting

      public void Set_Conditional_Formatting(string CellReferenceRange, ConditionalFormattingExpressionCondition Condition, string FormatValue, Color? Backcolour, Color? Forecolour, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.Range[CellReferenceRange];
            ExpressionConditionalFormatting cfRule = workSheet.ConditionalFormattings.AddExpressionConditionalFormatting(range, Condition, FormatValue);

            if (Backcolour.HasValue)
               cfRule.Formatting.Fill.BackgroundColor = Backcolour.Value;

            if (Forecolour.HasValue)
               cfRule.Formatting.Font.Color = Forecolour.Value;
         }
      }

        public void Set_Colour_Gradient_Formatting(string CellReferenceRange, string FormatValue, Color? Backcolour, Color? Forecolour, int SheetIndex)
        {
            Worksheet workSheet = workbook.Worksheets[SheetIndex];
            if (workSheet == null)
                throw new Exception("Unable to locate Sheet Index " + SheetIndex);
            else
            {
                ConditionalFormattingCollection conditionalFormattings = workSheet.ConditionalFormattings;
                // Set the minimum threshold to the lowest value in the range of cells.
                ConditionalFormattingValue minPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                ConditionalFormattingValue midPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.Percentile, "50");
                // Set the maximum threshold to the highest value in the range of cells.
                ConditionalFormattingValue maxPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Create the two-color scale rule to differentiate low and high values in cells C2 through D15. Blue represents the lower values and yellow represents the higher values. 
                ColorScale3ConditionalFormatting cfRule = conditionalFormattings.AddColorScale3ConditionalFormatting(workSheet.Range[CellReferenceRange], minPoint, Color.FromArgb(248, 105, 107),midPoint, Color.FromArgb(252, 252, 255), maxPoint, Color.FromArgb(99, 190, 123));
            }
        }


      #endregion
      #region Merge Cells

      public void Merge_Cells(string CellReferenceRange, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.Range[CellReferenceRange];
            range.Merge();
         }
      }

      #endregion
      #region Set Border

      public void Set_OutsideBorders(string CellReferenceRange, System.Drawing.Color? BorderColour, int SheetIndex, BorderLineStyle BorderStyle = BorderLineStyle.Medium)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.Range[CellReferenceRange];
            if (!BorderColour.HasValue)
               BorderColour = System.Drawing.Color.Black;

            range.Borders.SetOutsideBorders(BorderColour.Value, BorderStyle);
         }
      }

      public void Set_AllBorders(string CellReferenceRange, System.Drawing.Color? BorderColour, int SheetIndex)
      {
         Set_AllBorders(CellReferenceRange, BorderColour, BorderLineStyle.Medium, SheetIndex);
      }

      public void Set_AllBorders(string CellReferenceRange, System.Drawing.Color? BorderColour, BorderLineStyle bStyle, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            CellRange range = workSheet.Range[CellReferenceRange];
            if (!BorderColour.HasValue)
               BorderColour = System.Drawing.Color.Black;

            range.Borders.SetAllBorders(BorderColour.Value, bStyle);
         }
      }

      #endregion
      #region Wait Screens

      public void Show_Wait()
      {
         SplashScreenManager.ShowForm(typeof(frmWait));
      }
      public void Hide_Wait()
      {
         if (SplashScreenManager.Default != null && SplashScreenManager.Default.IsSplashFormVisible)
            SplashScreenManager.CloseForm();
      }

      #endregion
      #region Sheet Functions

      public void Set_Sheet_Name(int SheetIndex, string NewSheetName)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            if (NewSheetName.Length > 31)
               NewSheetName = NewSheetName.Substring(0, 28) + "...";

            workSheet.Name = NewSheetName;
         }
      }

      public void Set_Active_Sheet(int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];

         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
            workbook.Worksheets.ActiveWorksheet = workSheet;
      }


      #endregion
      #region Hide Columns

      public void Hide_Columns(int StartIndex, int EndIndex, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
            workSheet.Columns.Hide(StartIndex, EndIndex);
      }

      #endregion
      #region Images
      public void Insert_Image(string ImageName, string CellReference, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            if (File.Exists(ImageName))
            {
               //Image image = Image.FromFile(ImageName);
               //Stream imageStream = ImageToStream(image, Format);
               //SpreadsheetImageSource imageSource = SpreadsheetImageSource.FromStream(imageStream);
               workbook.BeginUpdate();
               try
               {
                  CellRange range = workSheet.Range[CellReference];
                  workSheet.Pictures.AddPicture(ImageName, range, true);
               }
               finally
               {
                  workbook.EndUpdate();
               }
            }
         }
      }

      public void Insert_Image(string ImageName, float x, float y, float Height, float Width, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            if (File.Exists(ImageName))
            {
               //Image image = Image.FromFile(ImageName);
               //Stream imageStream = ImageToStream(image, Format);
               //SpreadsheetImageSource imageSource = SpreadsheetImageSource.FromStream(imageStream);
               workbook.BeginUpdate();
               try
               {
                  workSheet.Pictures.AddPicture(ImageName, x, y, Width, Height, true);
               }
               finally
               {
                  workbook.EndUpdate();
               }
            }
         }
      }

      public void Insert_Image(Image mImage, float x, float y, float Height, float Width, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
         {
            workbook.BeginUpdate();
            try
            {
               workSheet.Pictures.AddPicture(mImage, x, y, Width, Height, true);
            }
            finally
            {
               workbook.EndUpdate();
            }
         }
      }

      //private Stream ImageToStream(this Image image, ImageFormat format)
      //{
      //   var stream = new System.IO.MemoryStream();
      //   image.Save(stream, format);
      //   stream.Position = 0;
      //   return stream;
      //}

      #endregion
      #region Row Height

      public void Set_Row_Height(int RowNumber, double RowHeight, int SheetIndex)
      {
         Worksheet workSheet = workbook.Worksheets[SheetIndex];
         if (workSheet == null)
            throw new Exception("Unable to locate Sheet Index " + SheetIndex);
         else
            workSheet.Rows[RowNumber].RowHeight = RowHeight;
      }

      #endregion
   }
}
