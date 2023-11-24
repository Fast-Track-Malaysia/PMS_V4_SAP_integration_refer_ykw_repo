using System;
using System.Collections.Generic;
using System.Text;

namespace PMS_V4_SAP_integration.Models
{
    public class CommissionList
    {
        public decimal UnitPrice { get; set; }
        public int Quantity { get; set; }
        public string ProjectCode { get; set; }
        public string VatGroup { get; set; }
        public string U_FRef { get; set; }
        public string ItemDescription { get; set; }


    }
}

