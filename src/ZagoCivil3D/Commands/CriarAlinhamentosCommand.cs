// CriarAlinhamentosCommand.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ZagoCivil3D.Commands
{
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
            List<string> estilosAlinhamento = ObterNomesEstilosAlinhamento(civilDoc, db);
            List<string> conjuntosRotulos = ObterNomesConjuntosRotulosAlinhamento(civilDoc, db);

            var tela = new CriarAlinhamentosWindow(camadas, estilosAlinhamento, conjuntosRotulos);
            bool? resultadoDialogo = tela.ShowDialog();

            if (resultadoDialogo != true)
                return null;

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

        private static List<string> ObterNomesEstilosAlinhamento(CivilDocument civilDoc, Database db)
        {
            var nomes = new List<string>();

            using Transaction transacao = db.TransactionManager.StartTransaction();

            foreach (ObjectId id in civilDoc.Styles.AlignmentStyles)
            {
                if (id.IsNull || id.IsErased)
                    continue;

                try
                {
                    if (transacao.GetObject(id, OpenMode.ForRead) is AlignmentStyle estilo &&
                        !string.IsNullOrWhiteSpace(estilo.Name))
                    {
                        nomes.Add(estilo.Name);
                    }
                }
                catch
                {
                }
            }

            transacao.Commit();
            return nomes.Distinct().OrderBy(x => x).ToList();
        }

        private static List<string> ObterNomesConjuntosRotulosAlinhamento(CivilDocument civilDoc, Database db)
        {
            var nomes = new List<string>();

            using Transaction transacao = db.TransactionManager.StartTransaction();

            foreach (ObjectId id in civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles)
            {
                if (id.IsNull || id.IsErased)
                    continue;

                try
                {
                    if (transacao.GetObject(id, OpenMode.ForRead) is AlignmentLabelSetStyle conjunto &&
                        !string.IsNullOrWhiteSpace(conjunto.Name))
                    {
                        nomes.Add(conjunto.Name);
                    }
                }
                catch
                {
                }
            }

            transacao.Commit();
            return nomes.Distinct().OrderBy(x => x).ToList();
        }
    }
}