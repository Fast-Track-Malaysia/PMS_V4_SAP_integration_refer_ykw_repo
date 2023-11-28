using System;
using System.Collections.Generic;
using System.Text;

namespace PMS_V4_SAP_integration.Models
{
    public class CommissionList
    {
        public string AccountCode { get; set; }
        public string ItemDescription { get; set; }
        public string ProjectCode { get; set; }
        public double Quantity { get; set; }
        public string U_FRef { get; set; }
        public double UnitPrice { get; set; }
        public string VatGroup { get; set; }
    }
}

