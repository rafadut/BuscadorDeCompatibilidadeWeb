﻿using LinqToExcel.Attributes;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace BuscadorDeCompatibilidadeWeb.Models
{
    public class VagaModel
    {
        public List<Vaga> ListaVaga { get; set; }
    }

    public class Vaga : Conhecimento
    {
        public const string ID_DA_VAGA = "#";
        public const string TITULO_DA_VAGA = "Título da vaga";

        [ExcelColumn(ID_DA_VAGA), Display(Name = ID_DA_VAGA)]
        public string VagaID { get; set; }

        [ExcelColumn(TITULO_DA_VAGA), Display(Name = TITULO_DA_VAGA)]
        public string VagaNome { get; set; }
    }
}