using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using System.Collections.Generic;
using System.Linq;
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;
// Desambigua: Autodesk.AutoCAD.Runtime tambem define Exception.
using Exception = System.Exception;

namespace ZagoCivil3D.Commands;

/// <summary>
/// Comando que define o FlowPath dos catchments existentes a partir das
/// polilinhas desenhadas em uma layer especifica (replica a rotina Dynamo
/// "CRIAR TALVEGUE DOS CATCHMENTS A PARTIR DAS POLILINHAS").
/// Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class CriarTalveguesCatchmentsCommand
{
    private static CriarTalveguesCatchmentsWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_CRIAR_TALVEGUES_CATCHMENTS", CommandFlags.Session)]
    public void Executar()
    {
        Document? documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
        if (documento == null)
            return;

        Database banco = documento.Database;
        Editor editor = documento.Editor;

        // Se ja existe janela aberta, apenas reativa — evita duas instancias.
        if (m_janelaAtiva != null)
        {
            m_janelaAtiva.Activate();
            return;
        }

        try
        {
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando criacao de talvegues dos catchments.");

            List<string> camadas = ObterTodosNomesCamadas(banco);

            var janela = new CriarTalveguesCatchmentsWindow(camadas);
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
        catch (Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro ao abrir janela: {ex.Message}");
        }
    }

    /// <summary>
    /// Handler do botao Executar da janela. Bloqueia o documento antes de
    /// chamar o servico (a janela modeless nao tem document lock implicito).
    /// </summary>
    private static void AoConfirmarJanela(object? remetente, CriarTalveguesCatchmentsRequest request)
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

            CriarTalveguesCatchmentsResultado resultado =
                TalveguesCatchmentsService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO TALVEGUES CATCHMENTS =====");
            editor.WriteMessage($"\nLayer das polilinhas: {request.NomeCamadaPolilinhas}");
            editor.WriteMessage($"\nPolilinhas encontradas: {resultado.TotalPolilinhas}");
            editor.WriteMessage($"\nCatchments no desenho: {resultado.TotalCatchments}");
            editor.WriteMessage($"\nFlowPaths criados: {resultado.TotalFlowPathsCriados}");
            editor.WriteMessage($"\nFlowPaths substituidos: {resultado.TotalFlowPathsSubstituidos}");
            editor.WriteMessage($"\nCatchments ignorados (FlowPath preservado): {resultado.TotalCatchmentsIgnorados}");
            editor.WriteMessage($"\nPolilinhas sem catchment correspondente: {resultado.TotalPolilinhasSemCatchment}");

            if (resultado.NomesCatchmentsAtualizados.Count > 0)
            {
                editor.WriteMessage("\nCatchments atualizados:");
                foreach (string nome in resultado.NomesCatchmentsAtualizados)
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

            int totalAtualizados = resultado.TotalFlowPathsCriados + resultado.TotalFlowPathsSubstituidos;
            string mensagemStatus = totalAtualizados > 0
                ? $"{totalAtualizados} catchment(s) atualizado(s)."
                : "Nenhum catchment atualizado.";

            if (resultado.TotalPolilinhasSemCatchment > 0)
                mensagemStatus += $" {resultado.TotalPolilinhasSemCatchment} polyline(s) sem catchment (revisar).";

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = totalAtualizados > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando de talvegues: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    /// <summary>Retorna todos os nomes de layer do desenho em ordem alfabetica.</summary>
    private static List<string> ObterTodosNomesCamadas(Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();
        var tabelaCamadas = (LayerTable)transacao.GetObject(db.LayerTableId, OpenMode.ForRead);

        foreach (ObjectId idCamada in tabelaCamadas)
        {
            if (transacao.GetObject(idCamada, OpenMode.ForRead) is LayerTableRecord registroCamada)
                nomes.Add(registroCamada.Name);
        }

        transacao.Commit();
        return nomes.OrderBy(x => x).ToList();
    }
}
