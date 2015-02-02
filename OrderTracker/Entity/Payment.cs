using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderTracker.Entity
{
    public class Payment
    {
        public String Category { get; set; }
        public String SKU_Code { get; set; }
        public String SuborderID { get; set; }
        public String CustomerName { get; set; }
        public String Shipped_Return_Date { get; set; }
        public String Delivered_Date { get; set; }
        public String Shipping_method_code { get; set; }
        public String Courier { get; set; }
        public String AWB_Number { get; set; }
        public decimal Amount { get; set; }
        public String RefNo { get; set; }
        public String Other_Applications { get; set; }
    }
}
