using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.Models
{
   public class CellModel
   {
      public int RowIndex { get; set; }
      public string RowReference { get; set; }
      public int ColumnIndex { get; set; }
      public string ColumnReference { get; set; }
      public int SheetIndex { get; set; }
      public string SheetName { get; set; }

      public Color Backcolour { get; set; }
      public Color Forecolour { get; set; }
   }
}
