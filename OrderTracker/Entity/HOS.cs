using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderTracker.Entity
{
   public class HOS
    {
       public String SubOrderID { get; set; }
       public String Sku { get; set; }
       public String Supc { get; set; }
       public String AWB { get; set; }
       public String Ref { get; set; }
       public DateTime CreationDate { get; set; }
       public String HosNo { get; set; }
       public String HosDate { get; set; }
    }
   public class HOSDetails
   {
       public String HosNo { get; set; }
       public String HosDate { get; set; }
       public int index { get; set; }
   }
   public class RefDetails
   {
     
       public String Other { get; set; }
       public int index { get; set; }
   }
}
