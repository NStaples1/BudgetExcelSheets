using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BudgetExcelSheets.Models
{
    internal class PriorYearSalesModel
    {
        public string  AccountRef { get; set; }
        public string Name { get; set; }
        public double? LineCostPrice { get; set; }
        public double? LineSalePrice { get; set; }
        public double? LineUnitWeight { get; set; }
        public int? InvoiceMonth { get; set; }

        public bool Deleted { get; set; }
    }
}
