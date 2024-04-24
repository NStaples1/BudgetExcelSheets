using DevExpress.Pdf;
using DevExpress.XtraPdfViewer;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DXTools
{
   public class PDFDocument : IDisposable
   {
      //private PdfDocumentProcessor pdfDocumentProcessor = new PdfDocumentProcessor();
      private const string ModuleName = "DXTools.Classes.PDFDocument";
      private PdfViewer Viewer = new PdfViewer();

      public enum PageRotationType : int
      {
         None = 0,
         Left = 270,
         right = 90,
         UpsideDown = 180
      }

      public enum PageOrientation
      {
         Portrait,
         Landscape,
         Unknown
      }

      public void Dispose()
      {
         if (Viewer != null)
         {
            Viewer.Dispose();
            Viewer = null;
         }

         //if (pdfDocumentProcessor != null)
         //{
         //   pdfDocumentProcessor.Dispose();
         //   pdfDocumentProcessor = null;
         //}
      }

      public bool Load(string Filename, bool ShowErrorMessage = true)
      {
         try
         {
            if (Viewer != null)
            {
               Viewer.LoadDocument(Filename);
               return true;
            }
            else
               return false;

            //if (pdfDocumentProcessor != null)
            //{
            //   pdfDocumentProcessor.LoadDocument(Filename);
            //   return true;
            //}
            //else
            //   return false;
         }
         catch (Exception ex)
         {
            if (ShowErrorMessage)
               ProcessError.Show(ModuleName, "Load", ex, new List<string>() { Filename });
            return false;
         }
      }

      public bool Load(System.IO.Stream Filestream, bool ShowErrorMessage = true)
      {
         try
         {
            if (Viewer != null)
            {
               Viewer.LoadDocument(Filestream);
               return true;
            }
            else
               return false;

            //if (pdfDocumentProcessor != null)
            //{
            //   pdfDocumentProcessor.LoadDocument(Filestream);
            //   return true;
            //}
            //else
            //   return false;
         }
         catch (Exception ex)
         {
            if (ShowErrorMessage)
               ProcessError.Show(ModuleName, "Load", ex);
            return false;
         }
      }

      public void RotatePage(PageRotationType pageRotationType)
      {
         if (Viewer != null)
         {
            int angle = (int)pageRotationType;
            Viewer.RotationAngle = angle;
         }


         //if (pdfDocumentProcessor != null)
         //{
         //   int angle = (int)pageRotationType;
         //   foreach (PdfPage page in pdfDocumentProcessor.Document.Pages)
         //   {
         //      page.Rotate = angle;
         //   }
         //}
      }

      public void Print(PrintDialog printdialog)
      {
         if (Viewer != null)
         {
            PdfPrinterSettings pdfPrinterSettings = new PdfPrinterSettings(printdialog.PrinterSettings);
            pdfPrinterSettings.PageOrientation = PdfPrintPageOrientation.Portrait;
            pdfPrinterSettings.ScaleMode = PdfPrintScaleMode.Fit;

            pdfPrinterSettings.PrintingDpi = 300;              // 200 was unreadable so its finding a balance 300 is readable but slightly slow on this large document but its going to have to do.
            pdfPrinterSettings.EnableLegacyPrinting = true;    // Had an issue with out of memory issues with a stannah pdf - this cured it

            Viewer.Print(pdfPrinterSettings);
         }
      }

      public PageOrientation GetPageOrientation()
      {
         if (Viewer != null)
         {
            SizeF FirstPageSize = Viewer.GetPageSize(1);
            if (FirstPageSize.Width > FirstPageSize.Height)
               return PageOrientation.Landscape;
            else
               return PageOrientation.Portrait;
         }
         else
            return PageOrientation.Unknown;


         //if (pdfDocumentProcessor != null)
         //{
         //   PdfPage FirstPage = pdfDocumentProcessor.Document.Pages[0];
         //   if (FirstPage != null)
         //   {
         //      // Check if the width is greater than the height to find the orientation
         //      PdfRectangle cropBox = FirstPage.CropBox;
         //      float cropBoxWidth = (float)cropBox.Width;
         //      float cropBoxHeight = (float)cropBox.Height;

         //      if (cropBoxWidth > cropBoxHeight)
         //         return PageOrientation.Landscape;
         //      else if (cropBoxWidth == cropBoxHeight)
         //         return PageOrientation.Unknown;
         //      else
         //         return PageOrientation.Portrait;
         //   }
         //   else
         //      return PageOrientation.Unknown;
         //}
         //else
         //   return PageOrientation.Unknown;
      }
   }
}
