using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderTracker.Entity
{
   public class OrderDetailsEntity
    {
       public String SuborderId { get; set; }
       public DateTime CreationDate { get; set; }
       public String Status { get; set; }
       public String Remark { get; set; }
       public decimal Amount { get; set; }
    }
}
