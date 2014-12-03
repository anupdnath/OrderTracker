using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderTracker.Entity
{
    public class UserLogin
    {
       public String UserName { get; set; }
       public String UserPass { get; set; }      
       public String UserType { get; set; }
       public DateTime CreationDate { get; set; }
       public DateTime UpdationDate { get; set; }
       public Boolean Active { get; set; }
       public UserLogin()
       {
           Active = true;
       }
    }
}
