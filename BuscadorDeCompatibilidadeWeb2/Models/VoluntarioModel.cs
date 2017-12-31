using LinqToExcel.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace BuscadorDeCompatibilidadeWeb.Models
{
    public class VoluntarioModel : Conhecimento
    {
        public const string NOME_COMPLETO = "Nome completo";
        public const string GENERO = "Gênero";
        public const string ENDERECO_COMPLETO = "Endereço completo";
        public const string TEL_RESIDENCIAL = "Telefone residencial";
        public const string CELULAR = "Telefone celular";
        public const string ESCOLARIDADE = "Nível de Escolaridade";
        public const string PROFISSAO = "Profissão";
        public const string AREA = "Área em que deseja atuar";
        public const string POSSUI_EXPERIENCIA = "Possui experiência em projetos sociais?";
        public const string QUAIS = "Quais?";
        public const string ATE = "até";
        public const string DISPONIBILIDADE = "Disponibilidade de dias";
        public const string MANHA = "Manhã";

        [ExcelColumn(NOME_COMPLETO), Display(Name = NOME_COMPLETO)]
        public string NomeCompleto { get; set; }
        [ExcelColumn(GENERO), Display(Name = GENERO)]
        public string Genero { get; set; }
        public string RG { get; set; }
        public string CPF { get; set; }
        [ExcelColumn(ENDERECO_COMPLETO), Display(Name = ENDERECO_COMPLETO)]
        public string EnderecoCompleto { get; set; }
        public string Bairro { get; set; }
        public string CEP { get; set; }
        public string Cidade { get; set; }
        public string Estado { get; set; }
        [ExcelColumn(TEL_RESIDENCIAL), Display(Name = TEL_RESIDENCIAL)]
        public string TelefoneResidencial { get; set; }
        [ExcelColumn(CELULAR), Display(Name = CELULAR)]
        public string TelefoneCelular { get; set; }
        public string Email { get; set; }
        [ExcelColumn(ESCOLARIDADE), Display(Name = ESCOLARIDADE)]
        public string NivelDeEscolaridade { get; set; }
        [ExcelColumn(PROFISSAO), Display(Name = PROFISSAO)]
        public string Profissao { get; set; }
        [ExcelColumn(AREA), Display(Name = AREA)]
        public string AreaEmQueDesejaAtuar { get; set; }
        public string Other { get; set; }
        [ExcelColumn(POSSUI_EXPERIENCIA), Display(Name = POSSUI_EXPERIENCIA)]
        public string PossuiExperienciaEmProjetosSociais { get; set; }
        [ExcelColumn(QUAIS), Display(Name = QUAIS)]
        public string Quais { get; set; }
        public string De { get; set; }
        [ExcelColumn(ATE), Display(Name = ATE)]
        public string Ate { get; set; }
        [ExcelColumn(DISPONIBILIDADE), Display(Name = DISPONIBILIDADE)]
        public string DisponibilidadeDeDias { get; set; }
        [ExcelColumn(MANHA), Display(Name = MANHA)]
        public string Manha { get; set; }
        public string Tarde { get; set; }
        public string Compatibilidade { get; set; }
    }
}