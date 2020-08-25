using ClosedXML.Excel;
using ReadExcelRedeDeliveryAPI.Models;
using ReadExcelRedeDeliveryAPI.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace ReadExcelRedeDeliveryAPI.Controllers
{
    public class ReadExcelController : ApiController
    {
        private readonly IReadExcelService _readExcelService;

        public ReadExcelController()
            //IReadExcelService readExcelService)
        {
           // this._readExcelService = readExcelService;
        }

        [Route("uploadRedeDelivery")]
        [HttpPost]
        public async Task<HttpResponseMessage> PostCarregarPlanilha()
        {
            string root = HttpContext.Current.Server.MapPath("~/App_Data");
            var provider = new MultipartFormDataStreamProvider(root);
            string fileLocalName = "";
            string path = System.AppDomain.CurrentDomain.BaseDirectory.ToString();

            var result = await Request.Content.ReadAsMultipartAsync(provider);
            foreach (MultipartFileData file in provider.FileData)
            {
                fileLocalName = file.LocalFileName;
            }

            File.WriteAllBytes($@"{path}\App_Data\saveCarregarRedeDelivery.xlsx", File.ReadAllBytes(fileLocalName));

            var fileName = $@"{path}\App_Data\saveCarregarRedeDelivery.xlsx";

            #region Tratativa do Excel
            using (var wb = new XLWorkbook(fileName))
            {

                var listaEstablecimentos = new List<Estabelecimento>();

                var linha = 2;

                var planilha = wb.Worksheet(1).RowsUsed();

                foreach (var dataRow in planilha)
                {
                    var nome = dataRow.Cell("A").Value.ToString();

                    if (string.IsNullOrEmpty(nome))
                    {
                        //tratativa de erro
                    }

                    var estabelecimento = new Estabelecimento()
                    {
                        NomeEstabelecimento = nome,
                        TipoCartao = dataRow.Cell("B").Value.ToString(),
                        Categoria = dataRow.Cell("C" ).Value.ToString(),
                        Estado = dataRow.Cell("D").Value.ToString(),
                        Cidade = dataRow.Cell("E").Value.ToString(),
                        Bairro = dataRow.Cell("F").Value.ToString(),
                        Endereco = dataRow.Cell("G").Value.ToString(),
                        Numero = dataRow.Cell("H").Value.ToString(),
                        Complemento = dataRow.Cell("I").Value.ToString(),
                        Whatsapp = dataRow.Cell("J").Value.ToString(),
                        Telefone = dataRow.Cell("K").Value.ToString(),
                        Site = dataRow.Cell("L").Value.ToString(),
                        HorarioFuncionamento = dataRow.Cell("M" ).Value.ToString()
                    };

                    estabelecimento.Delivery = dataRow.Cell("N").Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "UBER EATS|";
                    estabelecimento.Delivery = dataRow.Cell("O").Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "RAPPI|";
                    estabelecimento.Delivery = dataRow.Cell("P").Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "IFOOD|";
                    estabelecimento.Delivery = dataRow.Cell("Q").Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "DELIVERY|";
                    estabelecimento.Delivery = dataRow.Cell("R").Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "RETIRADA|";
                }

            }
            #endregion



            File.Delete($@"{path}\App_Data\saveCarregarRedeDelivery.xlsx");

            return new HttpResponseMessage();
        }
    }
}
