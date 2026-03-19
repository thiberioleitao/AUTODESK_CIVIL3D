// CriarAlinhamentosCommand.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ZagoCivil3D.Commands
{
    /// <summary>
    /// Comando principal que orquestra: coleta de parâmetros, execução do serviço
    /// e apresentação do resumo ao usuário na linha de comando.
    /// </summary>
    public class CriarAlinhamentosCommand
    {
        [CommandMethod("ZAGO_CRIAR_ALINHAMENTOS_POR_POLILINHA")]
        public void Executar()
        {
            var documento = Application.DocumentManager.MdiActiveDocument;
            if (documento == null)
                return;

            var banco = documento.Database;
            var editor = documento.Editor;
            CivilDocument documentoCivil = CivilApplication.ActiveDocument;

            try
            {
                editor.WriteMessage("\n[ZagoCivil3D] Comando iniciado.");

                var requisicao = ColetarParametros(documentoCivil, banco, editor);
                if (requisicao == null)
                {
                    editor.WriteMessage("\n[ZagoCivil3D] Comando cancelado.");
                    return;
                }

                CriarAlinhamentosResultado resultado =
                    AlignmentCreationService.Executar(documentoCivil, banco, editor, requisicao);

                editor.WriteMessage("\n");
                editor.WriteMessage("\n===== RESUMO =====");
                editor.WriteMessage($"\nPolilinhas encontradas: {resultado.TotalPolilinhas}");
                editor.WriteMessage($"\nAlignments criados: {resultado.TotalCriados}");
                editor.WriteMessage($"\nSem zona: {resultado.TotalSemZona}");

                if (resultado.NomesCriados.Count > 0)
                {
                    editor.WriteMessage("\nNomes criados:");
                    foreach (string nome in resultado.NomesCriados)
                        editor.WriteMessage($"\n - {nome}");
                }

                if (resultado.MensagensErro.Count > 0)
                {
                    editor.WriteMessage("\nMensagens/erros:");
                    foreach (string msg in resultado.MensagensErro)
                        editor.WriteMessage($"\n - {msg}");
                }

                editor.WriteMessage("\n===== FIM =====");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando: {ex.Message}");
            }
        }

        private static CriarAlinhamentosRequest? ColetarParametros(
            CivilDocument civilDoc,
            Database db,
            Editor ed)
        {
            List<string> camadas = ObterTodosNomesCamadas(db);
            List<string> estilosAlinhamento = ObterNomesEstilosAlinhamento(civilDoc);
            List<string> conjuntosRotulos = ObterNomesConjuntosRotulosAlinhamento(civilDoc);

            var tela = new CriarAlinhamentosWindow(camadas, estilosAlinhamento, conjuntosRotulos);
            bool? resultadoDialogo = tela.ShowDialog();

            if (resultadoDialogo != true)
                return null;

            // A UI atualmente trabalha com uma única layer de origem/destino.
            // Mantemos o mesmo valor para preservar comportamento previsível.
            return new CriarAlinhamentosRequest
            {
                Prefixo = tela.Prefixo,
                IdentificadorZona = tela.IdentificadorZona,
                NomeCamadaOrigem = tela.NomeCamadaOrigem,
                NomeCamadaDestino = tela.NomeCamadaOrigem,
                NomeEstiloAlinhamento = tela.NomeEstiloAlinhamento,
                NomeConjuntoRotulosAlinhamento = tela.NomeConjuntoRotulosAlinhamento,
                NumeroInicial = tela.NumeroInicial,
                Incremento = tela.Incremento,
                ApagarPolilinhasOriginais = tela.ApagarPolilinhasOriginais
            };
        }

        private static string? SolicitarTexto(Editor editor, string mensagem, string? valorPadrao)
        {
            // Mantido para eventual fallback em modo somente linha de comando.
            var opcoes = new PromptStringOptions(
                $"{mensagem}{(string.IsNullOrWhiteSpace(valorPadrao) ? "" : $" <{valorPadrao}>")}: ")
            {
                AllowSpaces = true
            };

            PromptResult resultado = editor.GetString(opcoes);
            if (resultado.Status == PromptStatus.Cancel)
                return null;

            if (string.IsNullOrWhiteSpace(resultado.StringResult))
                return valorPadrao ?? string.Empty;

            return resultado.StringResult;
        }

        private static int? SolicitarInteiro(Editor editor, string mensagem, int valorPadrao)
        {
            // Mantido para eventual fallback em modo somente linha de comando.
            var opcoes = new PromptIntegerOptions($"{mensagem} <{valorPadrao}>: ")
            {
                DefaultValue = valorPadrao,
                UseDefaultValue = true,
                AllowNegative = false,
                AllowZero = false,
                AllowNone = true
            };

            PromptIntegerResult resultado = editor.GetInteger(opcoes);
            if (resultado.Status == PromptStatus.Cancel)
                return null;

            if (resultado.Status == PromptStatus.None)
                return valorPadrao;

            return resultado.Value;
        }

        private static bool? SolicitarSimNao(Editor editor, string mensagem, bool valorPadrao)
        {
            // Mantido para eventual fallback em modo somente linha de comando.
            const string Sim = "Sim";
            const string Nao = "Nao";
            string palavraPadrao = valorPadrao ? "Sim" : "Nao";
            var opcoes = new PromptKeywordOptions(
                $"{mensagem} [{nameof(Sim)}/{nameof(Nao)}] <{palavraPadrao}>: ",
                "Sim Nao")
            {
                AllowNone = true
            };

            PromptResult resultado = editor.GetKeywords(opcoes);
            if (resultado.Status == PromptStatus.Cancel)
                return null;

            if (resultado.Status == PromptStatus.None)
                return valorPadrao;

            return string.Equals(resultado.StringResult, "Sim", StringComparison.OrdinalIgnoreCase);
        }

        private static List<string> ObterTodosNomesCamadas(Database db)
        {
            var nomes = new List<string>();

            using Transaction transacao = db.TransactionManager.StartTransaction();
            var tabelaCamadas = (LayerTable)transacao.GetObject(db.LayerTableId, OpenMode.ForRead);

            foreach (ObjectId id in tabelaCamadas)
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is LayerTableRecord registroCamada)
                    nomes.Add(registroCamada.Name);
            }

            transacao.Commit();
            return nomes.OrderBy(x => x).ToList();
        }

        private static List<string> ObterNomesEstilosAlinhamento(CivilDocument civilDoc)
        {
            var nomes = new List<string>();

            using Transaction transacao = civilDoc.Database.TransactionManager.StartTransaction();
            foreach (ObjectId idEstilo in civilDoc.Styles.AlignmentStyles)
            {
                DBObject estilo = transacao.GetObject(idEstilo, OpenMode.ForRead);
                string? nomeEstilo = estilo.GetType().GetProperty("Name")?.GetValue(estilo)?.ToString();
                if (!string.IsNullOrWhiteSpace(nomeEstilo))
                    nomes.Add(nomeEstilo);
            }

            transacao.Commit();
            if (nomes.Count == 0)
            {
                nomes.AddRange(civilDoc.Styles.AlignmentStyles
                    .Cast<ObjectId>()
                    .Select(x => x.ToString()));
            }

            return nomes.Distinct().OrderBy(x => x).ToList();
        }

        private static List<string> ObterNomesConjuntosRotulosAlinhamento(CivilDocument civilDoc)
        {
            var nomes = new List<string>();

            using Transaction transacao = civilDoc.Database.TransactionManager.StartTransaction();
            foreach (ObjectId idConjunto in civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles)
            {
                DBObject conjunto = transacao.GetObject(idConjunto, OpenMode.ForRead);
                string? nomeConjunto = conjunto.GetType().GetProperty("Name")?.GetValue(conjunto)?.ToString();
                if (!string.IsNullOrWhiteSpace(nomeConjunto))
                    nomes.Add(nomeConjunto);
            }

            transacao.Commit();
            if (nomes.Count == 0)
            {
                nomes.AddRange(civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles
                    .Cast<ObjectId>()
                    .Select(x => x.ToString()));
            }

            return nomes.Distinct().OrderBy(x => x).ToList();
        }
    }
}