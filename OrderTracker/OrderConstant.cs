using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrderTracker
{
    public sealed class OrderConstant
    {
        #region [Transection Comment]
         public const string Packed = "Order has been packed";
         public const string Shipped = "Item has been dispatched";
         public const string Returned = "Item has been returned";
         public const string PaymentReceived = "Payment Received";
         public const string Penalty = "Penalty Applied";
         public const string PenaltyApproved = "Penalty Approved";
        #endregion
    }
}
