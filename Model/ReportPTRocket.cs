using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePTGPS.Model
{
     public class ReportPTRocket
    {
        public int PaymentServiceID { get; set; }
        public string RegDate { get; set; }
        public string ServiceProviderCategory { get; set; }
        public int SocioID { get; set; }
        public string NameServiceProvider { get; set; }
        public int PersonID { get; set; }
        public string ReceiverName { get; set; }
        public decimal Amount { get; set; }
        public decimal Comission { get; set; }
        public decimal TotalWithComission { get; set; }
        public int Status { get; set; }
        public Int64 Receiver { get; set; }
        public string Identifier { get; set; }
        public string CreditCardNumber { get; set; }
    }
}
