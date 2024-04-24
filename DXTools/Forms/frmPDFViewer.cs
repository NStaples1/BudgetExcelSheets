using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DXTools.Forms
{
   public partial class frmPDFViewer : Form
   {
      public frmPDFViewer()
      {
         InitializeComponent();
      }

      public void Load_PDF(string documentLocation)
      {
         try
         {
            pdfViewer1.LoadDocument(documentLocation);

            // Check if we need to rotate the document
            SizeF FirstPageSize = pdfViewer1.GetPageSize(1);
            if (FirstPageSize.Width > FirstPageSize.Height)
               pdfViewer1.RotationAngle = 270;

            pdfViewer1.ZoomMode = DevExpress.XtraPdfViewer.PdfZoomMode.PageLevel;

            this.BringToFront();
            this.Show();
         }
         catch(Exception ex)
         {
            ProcessError.Show(this, "Load_PDF", ex, new List<string>() { documentLocation });
         }
      }
   }
}
