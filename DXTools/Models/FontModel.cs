using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.Models
{
   public class FontModel
   {
      public string FontName { get; set; }
      public int FontSize { get; set; }
      public Color FontColour { get; set; }
      public bool Bold { get; set; }
      public bool Underline { get; set; }

      public FontModel()
      {
         FontName = "Calibri";
         FontSize = 10;
         FontColour = Color.Black;
         Bold = false;
         Underline = false;
      }
   }
}
