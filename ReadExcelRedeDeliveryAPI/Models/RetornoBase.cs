using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelRedeDeliveryAPI.Models
{
    public class RetornoBase
    {
        public bool Sucesso { get; set; }
        public string Mensagem { get; set; }
        public object Content { get; set; }
    }
}