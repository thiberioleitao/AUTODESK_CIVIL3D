using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ZagoCivil3D.Commands;

[SupportedOSPlatform("windows")]
/// <summary>
/// Comando que exporta um CSV com dados de bacias e subbacias. Porta da rotina
/// Dynamo "EXPORTAR CSV BACIAS E SUBBACIAS - ID, AREA, TALVEGUE, DECLIVIDADE E
/// ID_JUSANTE". Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class ExportarCsvBaciasSubbaciasCommand
{
    private static ExportarCsvBaciasSubbaciasWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_EXPORTAR_CSV_BACIAS_SUBBACIAS", CommandFlags.Session)]
    public void Executar()
    {
        Document? documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
        if (documento == null)
            return;

        Editor editor = documento.Editor;

        if (m_janelaAtiva != null)
        {
            m_janelaAtiva.Activate();
            return;
        }

        try
        {
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando exportacao CSV de bacias e subbacias.");

            Database banco = documento.Database;
            CivilDocument documentoCivil = CivilApplication.ActiveDocument;

            List<string> layers = ObterNomesCamadas(banco);
            List<string> superficies = ObterNomesSuperficies(documentoCivil, banco);

            var janela = new ExportarCsvBaciasSubbaciasWindow(layers, superficies);
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

    private static void AoConfirmarJanela(object? remetente, ExportarCsvBaciasSubbaciasRequest request)
    {
        Document? documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
        if (documento == null)
            return;

        Database banco = documento.Database;
        Editor editor = documento.Editor;
        CivilDocument documentoCivil = CivilApplication.ActiveDocument;

        try
        {
            // Comando apenas le entidades e escreve arquivo — mesmo assim
            // bloqueamos o documento durante a leitura para evitar que
            // outro comando modifique o desenho no meio da coleta.
            using DocumentLock bloqueio = documento.LockDocument();

            ExportarCsvBaciasSubbaciasResultado resultado =
                ExportarCsvBaciasSubbaciasService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO EXPORTAR CSV BACIAS/SUBBACIAS =====");
            editor.WriteMessage($"\nArquivo: {(string.IsNullOrWhiteSpace(resultado.CaminhoCsvGerado) ? request.CaminhoCsv : resultado.CaminhoCsvGerado)}");
            editor.WriteMessage($"\nPrefixo das layers de hatches: {request.PrefixoLayersHatches}");
            editor.WriteMessage($"\nLayer dos MTexts: {request.NomeLayerMTextsSubbaciasId} (filtro: '{request.PrefixoTextoSubbacia}')");
            editor.WriteMessage($"\nLayer dos talvegues: {request.NomeLayerTalvegues}");
            editor.WriteMessage($"\nSuperficie: {(string.IsNullOrWhiteSpace(request.NomeSuperficie) ? "(nenhuma)" : request.NomeSuperficie)}");
            editor.WriteMessage($"\nRaio de busca de jusante: {request.RaioBuscaJusante} m");
            editor.WriteMessage($"\nHatches encontradas: {resultado.TotalHatches}");
            editor.WriteMessage($"\nMTexts encontrados: {resultado.TotalMTextsIds}");
            editor.WriteMessage($"\nTalvegues encontrados: {resultado.TotalTalvegues}");
            editor.WriteMessage($"\nRegioes de corredor: {resultado.TotalRegioesCorredores}");
            editor.WriteMessage($"\nLinhas no CSV: {resultado.TotalLinhasCsv}");
            editor.WriteMessage($"\nLinhas com aviso (REVISAR): {resultado.TotalLinhasComAviso}");

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

            string mensagemStatus;
            if (resultado.TotalLinhasCsv == 0)
            {
                mensagemStatus = resultado.MensagensErro.Count > 0
                    ? "Falha ao gerar o CSV. Consulte a linha de comando."
                    : "Nenhuma linha gerada.";
            }
            else
            {
                mensagemStatus = $"{resultado.TotalLinhasCsv} linha(s) exportada(s).";
                if (resultado.TotalLinhasComAviso > 0)
                    mensagemStatus += $" {resultado.TotalLinhasComAviso} com avisos (REVISAR).";
            }

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = resultado.TotalLinhasCsv > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
            m_janelaAtiva?.PopularResultados(resultado);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro na exportacao: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    /// <summary>Lista alfabetica das layers do desenho.</summary>
    private static List<string> ObterNomesCamadas(Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();
        var tabelaLayers = (LayerTable)transacao.GetObject(db.LayerTableId, OpenMode.ForRead);

        foreach (ObjectId id in tabelaLayers)
        {
            if (transacao.GetObject(id, OpenMode.ForRead) is LayerTableRecord registro
                && !string.IsNullOrWhiteSpace(registro.Name))
            {
                nomes.Add(registro.Name);
            }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }

    /// <summary>Lista alfabetica das superficies TIN do desenho.</summary>
    private static List<string> ObterNomesSuperficies(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.GetSurfaceIds())
        {
            if (transacao.GetObject(id, OpenMode.ForRead) is Autodesk.Civil.DatabaseServices.Surface sup
                && !string.IsNullOrWhiteSpace(sup.Name))
            {
                nomes.Add(sup.Name);
            }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }
}
