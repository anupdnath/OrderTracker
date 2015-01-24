using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderTracker.Entity
{
   public class OrderRef
    {
       public String OrderIdentifier { get; set; }
       public String Result { get; set; }      
    }
   public enum SaveType
   {
       None = 0,
       Mainfest = 1,
       HOS = 2,
       HOS1 = 3,
       Payment = 4
   }
}