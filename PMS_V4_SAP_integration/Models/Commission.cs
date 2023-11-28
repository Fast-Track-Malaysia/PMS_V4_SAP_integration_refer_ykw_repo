using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Text;

namespace PMS_V4_SAP_integration.Models
{
    public class Commission
    {
        public int sk_hdr { get; set; }
        public string CardCode { get; set; }
        public string NumAtCard { get; set; }
        public string Project { get; set; }
        public int HandWritten { get; set; }
        public BoDocumentTypes DocType { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DocDueDate { get; set; }
        public DateTime TaxDate { get; set; }
        public string Comments { get; set; }
        public int Series { get; set; }
        public BoYesNoEnum Handwritten { get; set; }

        public List<CommissionList> Lines { get; set; } = new List<CommissionList>();

    }
}
