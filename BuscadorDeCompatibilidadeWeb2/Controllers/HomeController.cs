using BuscadorDeCompatibilidadeWeb.Models;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace BuscadorDeCompatibilidadeWeb.Controllers
{
    public class HomeController : Controller
    {
        #region Constantes

        public const string NOME_ARQUIVO_EXCEL_VAGAS = "Vagas-report.xlsx";
        public const string NOME_ARQUIVO_EXCEL_VOLUNTARIOS = "Cadastro Estudante da Pátria-report.xlsx";
        public const string CONHECIMENTO_1 = "Word";
        public const string CONHECIMENTO_2 = "Excel";
        public const string CONHECIMENTO_3 = "PowerPoint";
        public const string CONHECIMENTO_4 = "Project";
        public const string CONHECIMENTO_5 = "Customer Relationship Management (CRM)";
        public const string CONHECIMENTO_6 = "Photoshop";
        public const string CONHECIMENTO_7 = "Corel";
        public const string CONHECIMENTO_8 = "Illustrator";
        public const string CONHECIMENTO_9 = "Fotografia";
        public const string CONHECIMENTO_10 = "InDesign";
        public const string NIVEL_0 = "Não sabe";
        public const string NIVEL_1 = "Sabe com ajuda";
        public const string NIVEL_2 = "Sabe com autonomia";
        public const string NIVEL_3 = "Sabe ensinar";

        #endregion

        #region Métodos

        public List<string> VerificarConhecimentosVaga(Vaga vagaSelecionada)
        {
            List<string> conhecimentosVaga = new List<string>();
            var props = vagaSelecionada.GetType().GetProperties();

            foreach (var prop in props)
            {
                string nomeConhecimentoVaga = prop.GetValue(vagaSelecionada).ToString();
                if (nomeConhecimentoVaga != "" && prop.Name != "VagaID" && prop.Name != "VagaNome")
                {
                    conhecimentosVaga.Add(nomeConhecimentoVaga);
                }
            }

            return conhecimentosVaga;
        }

        public static short ConverterParaPontuacao(string NivelConhecimento)
        {
            switch (NivelConhecimento)
            {
                case NIVEL_1: return 1;
                case NIVEL_2: return 2;
                case NIVEL_3: return 3;
                case NIVEL_0:
                default:
                    return 0;
            }
        }

        #endregion

        public ActionResult Index()
        {
            VagaModel model = new VagaModel();

            //Leitura da planilha

            var planilha = new ExcelQueryFactory(
                string.Format(@"{0}\Downloads\{1}",
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                NOME_ARQUIVO_EXCEL_VAGAS));

            var query =
                from c in planilha.Worksheet<Vaga>("results")
                where c.VagaID != ""
                select c;

            model.ListaVaga = query.ToList() as List<Vaga>;
            int quantidadeVagas = model.ListaVaga.Count();

            TempData["VagaModel"] = model;

            return View(model);
        }
        
        [HttpPost]
        public ActionResult SelecionarVaga(string ID)
        {
            VagaModel model = TempData["VagaModel"] as VagaModel;

            var vagaSelecionada = model.ListaVaga.First(m => m.VagaID.Equals(ID)) as Vaga;
            var conhecimentosVaga = VerificarConhecimentosVaga(vagaSelecionada);
            int quantidadeConhecimentosVaga = conhecimentosVaga.Count();
            int pontuacaoVaga = quantidadeConhecimentosVaga * 3;

            //Leitura da planilha

            var planilha = new ExcelQueryFactory(
                string.Format(@"{0}\Downloads\{1}",
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                NOME_ARQUIVO_EXCEL_VOLUNTARIOS));

            var query =
                from c in planilha.Worksheet<VoluntarioModel>("results")
                where c.Word != ""
                select c;

            var voluntarios = query.ToList() as List<VoluntarioModel>;
            int quantidadeVoluntarios = voluntarios.Count();

            for (int i = 0; i < quantidadeVoluntarios; i++)
            {
                int pontuacaoCandidato = 0;
                for (int j = 0; j < quantidadeConhecimentosVaga; j++)
                {
                    switch (conhecimentosVaga[j])
                    {
                        case CONHECIMENTO_1:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].Word);
                            break;
                        case CONHECIMENTO_2:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].Excel);
                            break;
                        case CONHECIMENTO_3:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].PowerPoint);
                            break;
                        case CONHECIMENTO_4:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].Project);
                            break;
                        case CONHECIMENTO_5:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].Crm);
                            break;
                        case CONHECIMENTO_6:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].Photoshop);
                            break;
                        case CONHECIMENTO_7:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].Corel);
                            break;
                        case CONHECIMENTO_8:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].Illustrator);
                            break;
                        case CONHECIMENTO_9:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].Fotografia);
                            break;
                        case CONHECIMENTO_10:
                        default:
                            pontuacaoCandidato += ConverterParaPontuacao(voluntarios[i].InDesign);
                            break;
                    }
                }

                //Cálculo de compatibilidade voluntário x vaga
                voluntarios[i].Compatibilidade = string.Concat(pontuacaoCandidato * 100 / pontuacaoVaga, "%");

                //Formatações
                if (voluntarios[i].Other != string.Empty)
                {
                    voluntarios[i].AreaEmQueDesejaAtuar = voluntarios[i].Other;
                }

                voluntarios[i].PossuiExperienciaEmProjetosSociais =
                    voluntarios[i].PossuiExperienciaEmProjetosSociais.Equals("1") ? "Sim" : "Não";

            }

            //Ordena por compatiblidade
            var voluntariosOrdenados = voluntarios.OrderByDescending(x => x.Compatibilidade).ToList();

            TempData["voluntariosOrdenados"] = voluntariosOrdenados;

            return RedirectToAction("Voluntarios");
        }

        public ActionResult Voluntarios()
        {
            List<VoluntarioModel> model = (List<VoluntarioModel>)TempData["voluntariosOrdenados"];

            return View(model);
        }

    }
}