using LinqToExcel.Attributes;
using System.ComponentModel.DataAnnotations;

namespace BuscadorDeCompatibilidadeWeb.Models
{
    public class Conhecimento
    {
        public const string CRM = "Customer Relationship Management (CRM)";

        public string Word { get; set; }
        public string Excel { get; set; }
        public string PowerPoint { get; set; }
        public string Project { get; set; }
        [ExcelColumn(CRM), Display(Name = CRM)]
        public string Crm { get; set; }
        public string Photoshop { get; set; }
        public string Corel { get; set; }
        public string Illustrator { get; set; }
        public string Fotografia { get; set; }
        public string InDesign { get; set; }
    }
}