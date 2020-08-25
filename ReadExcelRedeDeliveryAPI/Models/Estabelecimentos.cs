using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcelRedeDeliveryAPI.Models
{
    public class Estabelecimento
    {
        public string NomeEstabelecimento { get; set; }
        public string TipoCartao { get; set; }
        public string Categoria { get; set; }
        public string Estado { get; set; }
        public string Cidade { get; set; }
        public string Bairro { get; set; }
        public string Endereco { get; set; }
        public string Numero { get; set; }
        public string Complemento { get; set; }
        public string Whatsapp { get; set; }
        public string Telefone { get; set; }
        public string Site { get; set; }
        public string HorarioFuncionamento { get; set; }
        public string Delivery { get; set; }

    }
}