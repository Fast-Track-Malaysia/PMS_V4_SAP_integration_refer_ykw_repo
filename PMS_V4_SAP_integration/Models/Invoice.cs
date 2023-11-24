using System;
using System.Collections.Generic;
using System.Text;

namespace PMS_V4_SAP_integration.Models
{
    public class Invoice
    {
        public string CardCode { get; set; }
        public string NumAtCard { get; set; }
        public string Project { get; set; }
        public string DocDate { get; set; }
        public string DocDueDate { get; set; }
        public string TaxDate { get; set; }
        public string Comments { get; set; }
        public string Series { get; set; }

        public List<InvoiceList> Lines { get; set; } = new List<InvoiceList>();

    }
}
