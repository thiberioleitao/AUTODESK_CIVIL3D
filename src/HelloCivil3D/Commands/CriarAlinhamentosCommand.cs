// CriarAlinhamentosCommand.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using HelloCivil3D.Models;
using HelloCivil3D.Services;
using HelloCivil3D.Views;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HelloCivil3D.Commands
{
    /// <summary>
    /// Comando principal que orquestra: coleta de parâmetros, execução do serviço
    /// e apresentação do resumo ao usuário na linha de comando.
    /// </summary>
    public class CriarAlinhamentosCommand
    {
        [CommandMethod("PL_CRIAR_ALINHAMENTOS")]
        public void Executar()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
                return;

            var db = doc.Database;
            var ed = doc.Editor;
            CivilDocument civilDoc = CivilApplication.ActiveDocument;

            try
            {
                ed.WriteMessage("\n[HelloCivil3D] Comando iniciado.");

                var request = ColetarParametros(civilDoc, db, ed);
                if (request == null)
                {
                    ed.WriteMessage("\n[HelloCivil3D] Comando cancelado.");
                    return;
                }

                CriarAlinhamentosResultado resultado =
                    AlignmentCreationService.Executar(civilDoc, db, ed, request);

                ed.WriteMessage("\n");
                ed.WriteMessage("\n===== RESUMO =====");
                ed.WriteMessage($"\nPolilinhas encontradas: {resultado.TotalPolilinhas}");
                ed.WriteMessage($"\nAlignments criados: {resultado.TotalCriados}");
                ed.WriteMessage($"\nSem zona: {resultado.TotalSemZona}");

                if (resultado.NomesCriados.Count > 0)
                {
                    ed.WriteMessage("\nNomes criados:");
                    foreach (string nome in resultado.NomesCriados)
                        ed.WriteMessage($"\n - {nome}");
                }

                if (resultado.MensagensErro.Count > 0)
                {
                    ed.WriteMessage("\nMensagens/erros:");
                    foreach (string msg in resultado.MensagensErro)
                        ed.WriteMessage($"\n - {msg}");
                }

                ed.WriteMessage("\n===== FIM =====");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[HelloCivil3D] Erro no comando: {ex.Message}");
            }
        }

        private static CriarAlinhamentosRequest? ColetarParametros(
            CivilDocument civilDoc,
            Database db,
            Editor ed)
        {
            List<string> layers = GetAllLayerNames(db);
            List<string> sites = GetAllSiteNames(civilDoc, db);
            List<string> alignmentStyles = GetAlignmentStyleNames(civilDoc);
            List<string> labelSets = GetAlignmentLabelSetNames(civilDoc);

            var tela = new CriarAlinhamentosWindow(layers, sites, alignmentStyles, labelSets);
            bool? dialogResult = tela.ShowDialog();

            if (dialogResult != true)
                return null;

            // A UI atualmente trabalha com uma única layer de origem/destino.
            // Mantemos o mesmo valor para preservar comportamento previsível.
            return new CriarAlinhamentosRequest
            {
                Prefixo = tela.Prefixo,
                SourceLayerName = tela.SourceLayerName,
                DestinationLayerName = tela.SourceLayerName,
                SiteName = tela.SiteName,
                AlignmentStyleName = tela.AlignmentStyleName,
                AlignmentLabelSetName = tela.AlignmentLabelSetName,
                NumeroInicial = tela.NumeroInicial,
                Incremento = tela.Incremento,
                ApagarPolilinhasOriginais = tela.ApagarPolilinhasOriginais
            };
        }

        private static string? PromptForString(Editor ed, string message, string? defaultValue)
        {
            // Mantido para eventual fallback em modo somente linha de comando.
            var opts = new PromptStringOptions(
                $"{message}{(string.IsNullOrWhiteSpace(defaultValue) ? "" : $" <{defaultValue}>")}: ")
            {
                AllowSpaces = true
            };

            PromptResult res = ed.GetString(opts);
            if (res.Status == PromptStatus.Cancel)
                return null;

            if (string.IsNullOrWhiteSpace(res.StringResult))
                return defaultValue ?? string.Empty;

            return res.StringResult;
        }

        private static int? PromptForInt(Editor ed, string message, int defaultValue)
        {
            // Mantido para eventual fallback em modo somente linha de comando.
            var opts = new PromptIntegerOptions($"{message} <{defaultValue}>: ")
            {
                DefaultValue = defaultValue,
                UseDefaultValue = true,
                AllowNegative = false,
                AllowZero = false,
                AllowNone = true
            };

            PromptIntegerResult res = ed.GetInteger(opts);
            if (res.Status == PromptStatus.Cancel)
                return null;

            if (res.Status == PromptStatus.None)
                return defaultValue;

            return res.Value;
        }

        private static bool? PromptForYesNo(Editor ed, string message, bool defaultValue)
        {
            // Mantido para eventual fallback em modo somente linha de comando.
            const string Sim = "Sim";
            const string Nao = "Nao";
            string defaultKeyword = defaultValue ? "Sim" : "Nao";
            var opts = new PromptKeywordOptions(
                $"{message} [{nameof(Sim)}/{nameof(Nao)}] <{defaultKeyword}>: ",
                "Sim Nao")
            {
                AllowNone = true
            };

            PromptResult res = ed.GetKeywords(opts);
            if (res.Status == PromptStatus.Cancel)
                return null;

            if (res.Status == PromptStatus.None)
                return defaultValue;

            return string.Equals(res.StringResult, "Sim", StringComparison.OrdinalIgnoreCase);
        }

        private static List<string> GetAllLayerNames(Database db)
        {
            var names = new List<string>();

            using Transaction tr = db.TransactionManager.StartTransaction();
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            foreach (ObjectId id in lt)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is LayerTableRecord ltr)
                    names.Add(ltr.Name);
            }

            tr.Commit();
            return names.OrderBy(x => x).ToList();
        }

        private static List<string> GetAllSiteNames(CivilDocument civilDoc, Database db)
        {
            var names = new List<string>();

            using Transaction tr = db.TransactionManager.StartTransaction();
            foreach (ObjectId siteId in civilDoc.GetSiteIds())
            {
                if (tr.GetObject(siteId, OpenMode.ForRead) is Autodesk.Civil.DatabaseServices.Site site)
                    names.Add(site.Name);
            }

            tr.Commit();
            return names.OrderBy(x => x).ToList();
        }

        private static List<string> GetAlignmentStyleNames(CivilDocument civilDoc)
        {
            var names = new List<string>();

            try
            {
                for (int i = 0; i < civilDoc.Styles.AlignmentStyles.Count; i++)
                    names.Add(civilDoc.Styles.AlignmentStyles[i].ToString());
            }
            catch
            {
            }

            return names.Distinct().OrderBy(x => x).ToList();
        }

        private static List<string> GetAlignmentLabelSetNames(CivilDocument civilDoc)
        {
            var names = new List<string>();

            try
            {
                for (int i = 0; i < civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles.Count; i++)
                    names.Add(civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles[i].ToString());
            }
            catch
            {
            }

            return names.Distinct().OrderBy(x => x).ToList();
        }
    }
}