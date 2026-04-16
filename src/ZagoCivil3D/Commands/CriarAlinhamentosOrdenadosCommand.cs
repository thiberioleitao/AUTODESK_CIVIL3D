// CriarAlinhamentosOrdenadosCommand.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices.Styles;
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ZagoCivil3D.Commands
{
    /// <summary>
    /// Comando que cria alignments a partir de dois layers separados:
    /// primeiro as polilinhas horizontais ordenadas Norte→Sul,
    /// em seguida as verticais ordenadas Oeste→Leste,
    /// com numeração sequencial compartilhada.
    ///
    /// A janela é exibida em modo modeless (não bloqueia o Civil 3D),
    /// permitindo ao usuário navegar pelo desenho enquanto ela está aberta.
    /// </summary>
    public class CriarAlinhamentosOrdenadosCommand
    {
        // Referência estática da janela aberta para garantir que apenas uma
        // instância exista ao mesmo tempo. Em modeless precisamos manter o
        // controle vivo enquanto o usuário interage.
        private static CriarAlinhamentosOrdenadosWindow? m_janelaAtiva;

        [CommandMethod("ZAGO_CRIAR_ALINHAMENTOS_ORDENADOS", CommandFlags.Session)]
        public void Executar()
        {
            var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
            if (documento == null)
                return;

            var editor = documento.Editor;

            try
            {
                editor.WriteMessage("\n[ZagoCivil3D] Abrindo janela (ordenado horizontais N→S + verticais O→L).");

                if (m_janelaAtiva != null)
                {
                    m_janelaAtiva.Activate();
                    return;
                }

                var banco = documento.Database;
                CivilDocument documentoCivil = CivilApplication.ActiveDocument;

                List<string> camadas = ObterTodosNomesCamadas(banco);
                List<string> estilosAlinhamento = ObterNomesEstilosAlinhamento(documentoCivil, banco);
                List<string> conjuntosRotulos = ObterNomesConjuntosRotulosAlinhamento(documentoCivil, banco);

                var janela = new CriarAlinhamentosOrdenadosWindow(camadas, estilosAlinhamento, conjuntosRotulos);
                janela.ConfirmarClicado += AoConfirmarJanela;
                janela.Closed += (s, e) =>
                {
                    if (m_janelaAtiva != null)
                        m_janelaAtiva.ConfirmarClicado -= AoConfirmarJanela;
                    m_janelaAtiva = null;
                };

                m_janelaAtiva = janela;
                AplicacaoAutoCad.ShowModelessWindow(janela);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando: {ex.Message}");
            }
        }

        /// <summary>
        /// Handler disparado pela janela modeless quando o usuário clica em
        /// "Criar Alignments". Bloqueia o documento ativo e delega ao serviço.
        /// </summary>
        private static void AoConfirmarJanela(object? remetente, CriarAlinhamentosOrdenadosRequest requisicao)
        {
            var janela = remetente as CriarAlinhamentosOrdenadosWindow;
            var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
            if (documento == null)
            {
                janela?.DefinirStatus("Nenhum documento ativo.", sucesso: false);
                return;
            }

            var editor = documento.Editor;
            var banco = documento.Database;

            try
            {
                // Em comandos com CommandFlags.Session a janela não detém o lock
                // do documento. Precisamos adquirir explicitamente antes de
                // modificar entidades.
                using DocumentLock bloqueio = documento.LockDocument();

                CivilDocument documentoCivil = CivilApplication.ActiveDocument;

                editor.WriteMessage("\n[ZagoCivil3D] Executando criação de alignments…");

                CriarAlinhamentosResultado resultado =
                    AlignmentCreationService.ExecutarOrdenado(documentoCivil, banco, editor, requisicao);

                editor.WriteMessage("\n");
                editor.WriteMessage("\n===== RESUMO (ORDENADO) =====");
                editor.WriteMessage($"\nPolilinhas encontradas: {resultado.TotalPolilinhas}");
                editor.WriteMessage($"\nAlignments criados: {resultado.TotalCriados}");

                if (resultado.NomesCriados.Count > 0)
                {
                    string descricaoOrdem = requisicao.HorizontaisPrimeiro
                        ? "horizontais N→S, depois verticais O→L"
                        : "verticais O→L, depois horizontais N→S";
                    editor.WriteMessage($"\nNomes criados ({descricaoOrdem}):");
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

                string mensagemStatus =
                    $"Criação concluída: {resultado.TotalCriados}/{resultado.TotalPolilinhas} alignments."
                    + (resultado.MensagensErro.Count > 0
                        ? $" {resultado.MensagensErro.Count} aviso(s) — veja a linha de comando."
                        : string.Empty);

                bool sucesso = resultado.TotalCriados > 0 && resultado.MensagensErro.Count == 0;
                janela?.DefinirStatus(mensagemStatus, sucesso);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\n[ZagoCivil3D] Erro durante execução: {ex.Message}");
                janela?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
            }
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
