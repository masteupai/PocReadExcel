using ClosedXML.Excel;
using ReadExcelRedeDeliveryAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelRedeDeliveryAPI.Services.Interfaces
{
    interface IReadExcelService
    {
        Task<RetornoBase> ReadExcelAndInsertData(XLWorkbook wb);
    }
}
