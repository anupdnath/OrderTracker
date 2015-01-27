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
       public DateTime HSDate { get; set; }
       public int RefIndex { get; set; }
       public DateTime MainfeastDate { get; set; }
       public String Product { get; set; }
       public String RecDetails { get; set; }
       public String Weight { get; set; }
       public String MobileNo { get; set; }
       public HOS()
       {
           MainfeastDate = DateTime.Parse("1900-01-01");
           HSDate = DateTime.Parse("1900-01-01");
           Weight = "0";
           MobileNo = "";
           HosNo = "";
       }
      
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
