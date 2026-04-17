using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ZagoCivil3D.Commands;

/// <summary>
/// Comando que aplica um Alignment Label Set Style a todos os alinhamentos
/// do desenho. Replica a rotina Dynamo "MUDAR LABEL SET DOS ALINHAMENTOS".
/// Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class MudarLabelSetAlinhamentosCommand
{
    private static MudarLabelSetAlinhamentosWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_MUDAR_LABEL_SET_ALINHAMENTOS", CommandFlags.Session)]
    public void Executar()
    {
        Document? documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
        if (documento == null)
            return;

        Database banco = documento.Database;
        Editor editor = documento.Editor;
        CivilDocument documentoCivil = CivilApplication.ActiveDocument;

        // Se ja existe janela aberta, apenas reativa
        if (m_janelaAtiva != null)
        {
            m_janelaAtiva.Activate();
            return;
        }

        try
        {
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando troca de label set dos alinhamentos.");

            List<string> labelSets = ObterNomesLabelSetStyles(documentoCivil, banco);

            var janela = new MudarLabelSetAlinhamentosWindow(labelSets);

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
            editor.WriteMessage($"\n[ZagoCivil3D] Erro ao abrir janela: {ex.Message}");
        }
    }

    private static void AoConfirmarJanela(object? remetente, MudarLabelSetAlinhamentosRequest request)
    {
        Document? documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
        if (documento == null)
            return;

        Database banco = documento.Database;
        Editor editor = documento.Editor;
        CivilDocument documentoCivil = CivilApplication.ActiveDocument;

        try
        {
            using DocumentLock bloqueio = documento.LockDocument();

            MudarLabelSetAlinhamentosResultado resultado =
                MudarLabelSetAlinhamentosService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO MUDAR LABEL SET =====");
            editor.WriteMessage($"\nLabel Set aplicado: {request.NomeLabelSetStyle}");
            editor.WriteMessage($"\nApagar existentes: {(request.ApagarExistentes ? "sim" : "nao")}");
            editor.WriteMessage($"\nAlinhamentos encontrados: {resultado.TotalAlinhamentos}");
            editor.WriteMessage($"\nAplicados com sucesso: {resultado.TotalAplicadosComSucesso}");
            editor.WriteMessage($"\nLabel groups apagados: {resultado.TotalLabelGroupsApagados}");

            if (resultado.NomesProcessados.Count > 0)
            {
                editor.WriteMessage("\nAlinhamentos processados:");
                foreach (string nome in resultado.NomesProcessados)
                    editor.WriteMessage($"\n  - {nome}");
            }

            if (resultado.Logs.Count > 0)
            {
                editor.WriteMessage("\nLogs:");
                foreach (string log in resultado.Logs)
                    editor.WriteMessage($"\n  - {log}");
            }

            if (resultado.MensagensErro.Count > 0)
            {
                editor.WriteMessage("\nErros:");
                foreach (string erro in resultado.MensagensErro)
                    editor.WriteMessage($"\n  - {erro}");
            }

            editor.WriteMessage("\n===== FIM =====");

            // Status na janela
            string mensagemStatus = resultado.TotalAplicadosComSucesso > 0
                ? $"{resultado.TotalAplicadosComSucesso} alinhamento(s) processado(s) com sucesso."
                : "Nenhum alinhamento processado.";

            if (resultado.TotalLabelGroupsApagados > 0)
                mensagemStatus += $" {resultado.TotalLabelGroupsApagados} label group(s) apagado(s).";

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = resultado.TotalAplicadosComSucesso > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando de label set: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    /// <summary>
    /// Lista os nomes dos Alignment Label Set Styles disponiveis no desenho,
    /// em ordem alfabetica.
    /// </summary>
    private static List<string> ObterNomesLabelSetStyles(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is StyleBase estilo
                    && !string.IsNullOrWhiteSpace(estilo.Name))
                {
                    nomes.Add(estilo.Name);
                }
            }
            catch { }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }
}
