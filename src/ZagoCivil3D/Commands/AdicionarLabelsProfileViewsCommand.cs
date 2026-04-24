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
/// Comando que adiciona Profile Line Labels em todos os profile views do
/// desenho. Replica a rotina Dynamo "ADICIONAR LABELS AS PROFILE VIEWS":
/// para cada alinhamento, percorre seus profile views e, nos profiles cujo
/// nome contem o texto do filtro, cria um ProfileLineLabelGroup com o
/// Profile Line Label Style escolhido. Usa janela modeless para nao
/// bloquear o Civil 3D.
/// </summary>
public class AdicionarLabelsProfileViewsCommand
{
    private static AdicionarLabelsProfileViewsWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_ADICIONAR_LABELS_PROFILE_VIEWS", CommandFlags.Session)]
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
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando adicao de labels aos profile views.");

            List<string> estilos = ObterNomesProfileLineLabelStyles(documentoCivil, banco);

            var janela = new AdicionarLabelsProfileViewsWindow(estilos);

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

    private static void AoConfirmarJanela(object? remetente, AdicionarLabelsProfileViewsRequest request)
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

            AdicionarLabelsProfileViewsResultado resultado =
                AdicionarLabelsProfileViewsService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO ADICIONAR LABELS PROFILE VIEWS =====");
            editor.WriteMessage($"\nLabel Style aplicado: {request.NomeProfileLineLabelStyle}");
            editor.WriteMessage($"\nFiltro de nome (contem): '{request.FiltroNomeProfile}'");
            editor.WriteMessage($"\nAlinhamentos encontrados: {resultado.TotalAlinhamentos}");
            editor.WriteMessage($"\nProfile views encontrados: {resultado.TotalProfileViews}");
            editor.WriteMessage($"\nProfiles filtrados: {resultado.TotalProfilesFiltrados}");
            editor.WriteMessage($"\nLabels criados: {resultado.TotalLabelsCriados}");

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

            string mensagemStatus = resultado.TotalLabelsCriados > 0
                ? $"{resultado.TotalLabelsCriados} label(s) criado(s) com sucesso."
                : "Nenhum label criado.";

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = resultado.TotalLabelsCriados > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando de labels: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    /// <summary>
    /// Lista os nomes dos Profile Line Label Styles disponiveis no desenho,
    /// em ordem alfabetica.
    /// </summary>
    private static List<string> ObterNomesProfileLineLabelStyles(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.LabelStyles.ProfileLabelStyles.LineLabelStyles)
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
