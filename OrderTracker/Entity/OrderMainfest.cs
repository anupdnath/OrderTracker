using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderTracker.Entity
{
    public class OrderMainfest
    {

        public String Category { get; set; }
        public String Courier { get; set; }
        public String Product { get; set; }
        public String Reference_Code { get; set; }
        public String Suborder_Id { get; set; }
        public String SKU_Code { get; set; }
        public String AWB_Number { get; set; }
        public DateTime Order_Verified_Date { get; set; }
        public DateTime Order_Created_Date { get; set; }
        public String Customer_Name { get; set; }
        public String Shipping_Name { get; set; }
        public String City { get; set; }
        public String State { get; set; }
        public String PINCode { get; set; }
        public decimal Selling_Price { get; set; }
        public String IMEI_SERIAL { get; set; }
        public DateTime PromisedShipDate { get; set; }
        public decimal MRP { get; set; }
        public String InvoiceCode { get; set; }
        public DateTime CreationDate { get; set; }

    }
}
