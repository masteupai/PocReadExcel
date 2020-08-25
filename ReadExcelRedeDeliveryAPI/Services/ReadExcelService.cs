using ClosedXML.Excel;
using ReadExcelRedeDeliveryAPI.Models;
using ReadExcelRedeDeliveryAPI.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace ReadExcelRedeDeliveryAPI
{
    public class ReadExcelService : IReadExcelService
    {
        public async Task<RetornoBase> ReadExcelAndInsertData(XLWorkbook wb)
        {
            var retornoBase = new RetornoBase();
            var planilha = wb.Worksheet(1);

            var listaEstablecimentos = new List<Estabelecimento>();

            var linha = 2;
            while (true)
            {
                var nome = planilha.Cell("A" + linha.ToString()).Value.ToString();

                if (string.IsNullOrEmpty(nome))
                {
                    retornoBase.Sucesso = false;
                    retornoBase.Mensagem = "Nome do Estabelecimento não encontrado";
                    return retornoBase;
                }
                    

                var estabelecimento = new Estabelecimento()
                {
                    NomeEstabelecimento = nome,
                    TipoCartao = planilha.Cell("B" + linha.ToString()).Value.ToString(),
                    Categoria = planilha.Cell("C" + linha.ToString()).Value.ToString(),
                    Estado = planilha.Cell("D" + linha.ToString()).Value.ToString(), 
                    Cidade = planilha.Cell("E" + linha.ToString()).Value.ToString(), 
                    Bairro = planilha.Cell("F" + linha.ToString()).Value.ToString(), 
                    Endereco = planilha.Cell("G" + linha.ToString()).Value.ToString(),
                    Numero = planilha.Cell("H" + linha.ToString()).Value.ToString(),
                    Complemento = planilha.Cell("I" + linha.ToString()).Value.ToString(),
                    Whatsapp = planilha.Cell("J" + linha.ToString()).Value.ToString(),
                    Telefone = planilha.Cell("K" + linha.ToString()).Value.ToString(),
                    Site = planilha.Cell("L" + linha.ToString()).Value.ToString(),
                    HorarioFuncionamento = planilha.Cell("M" + linha.ToString()).Value.ToString()
                };

                estabelecimento.Delivery = planilha.Cell("N" + linha.ToString()).Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "|";
                estabelecimento.Delivery = planilha.Cell("O" + linha.ToString()).Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "|";
                estabelecimento.Delivery = planilha.Cell("P" + linha.ToString()).Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "|";
                estabelecimento.Delivery = planilha.Cell("Q" + linha.ToString()).Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "|";
                estabelecimento.Delivery = planilha.Cell("R" + linha.ToString()).Value.ToString().Equals("NÃO") ? estabelecimento.Delivery += "" : estabelecimento.Delivery += "|";


            }

            retornoBase.Content = listaEstablecimentos;
            return retornoBase;
        }
    }
}