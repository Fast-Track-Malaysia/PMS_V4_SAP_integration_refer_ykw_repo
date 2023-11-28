using System;
using System.Collections.Generic;
using System.Text;

namespace PMS_V4_SAP_integration.Models
{
    public class Invoice
    {
        public int sk_hdr { get; set; }
        public string CardCode { get; set; }
        public string NumAtCard { get; set; }
        public string Project { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DocDueDate { get; set; }
        public DateTime TaxDate { get; set; } //check date time format if it is matching with SAP
        public string Comments { get; set; }
        public int Series { get; set; }

        public List<InvoiceList> Lines { get; set; } = new List<InvoiceList>();

    }
}
