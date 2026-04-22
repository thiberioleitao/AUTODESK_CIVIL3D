using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
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
/// Comando que cria catchments a partir das hatches de uma layer.
/// Porta para C# da rotina Dynamo
/// "CRIAR CATCHMENT A PARTIR DAS HATCHS". Usa janela modeless para nao
/// bloquear o Civil 3D.
/// </summary>
public class CriarCatchmentsDeHatchsCommand
{
    private static CriarCatchmentsDeHatchsWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_CRIAR_CATCHMENTS_DE_HATCHS", CommandFlags.Session)]
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
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando criacao de catchments a partir de hatches.");

            Database banco = documento.Database;
            CivilDocument documentoCivil = CivilApplication.ActiveDocument;

            List<string> layers = ObterNomesCamadas(banco);
            List<string> superficies = ObterNomesSuperficies(documentoCivil, banco);
            List<string> gruposCatchment = ObterNomesGruposCatchment(banco);
            List<string> estilosCatchment = ObterNomesEstilosCatchment(documentoCivil, banco);

            var janela = new CriarCatchmentsDeHatchsWindow(layers, superficies, gruposCatchment, estilosCatchment);

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

    private static void AoConfirmarJanela(object? remetente, CriarCatchmentsDeHatchsRequest request)
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

            CriarCatchmentsDeHatchsResultado resultado =
                CriarCatchmentsDeHatchsService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO CRIAR CATCHMENTS =====");
            editor.WriteMessage($"\nLayer das hatches: {request.NomeLayerHatches}");
            editor.WriteMessage($"\nLayer dos MTexts (IDs): {request.NomeLayerMTextsIds}");
            editor.WriteMessage($"\nLayer dos talvegues: {(string.IsNullOrWhiteSpace(request.NomeLayerTalvegues) ? "(nenhuma)" : request.NomeLayerTalvegues)}");
            editor.WriteMessage($"\nCatchment Group: {request.NomeGrupoCatchment}");
            editor.WriteMessage($"\nPrefixo: {request.PrefixoBacia}");
            editor.WriteMessage($"\nSuperficie: {(string.IsNullOrWhiteSpace(request.NomeSuperficie) ? "(nenhuma)" : request.NomeSuperficie)}");
            editor.WriteMessage($"\nEstilo: {(string.IsNullOrWhiteSpace(request.NomeEstiloCatchment) ? "(padrao)" : request.NomeEstiloCatchment)}");
            editor.WriteMessage($"\nHatches encontradas: {resultado.TotalHatchesEncontradas}");
            editor.WriteMessage($"\nMTexts encontrados: {resultado.TotalMTextsEncontrados}");
            editor.WriteMessage($"\nTalvegues encontrados: {resultado.TotalTalveguesEncontrados}");
            editor.WriteMessage($"\nCatchments criados: {resultado.TotalCatchmentsCriados}");
            editor.WriteMessage($"\nCom flow path: {resultado.TotalComFlowPath}");
            if (resultado.TotalSubstituidos > 0)
                editor.WriteMessage($"\nSubstituidos: {resultado.TotalSubstituidos}");

            if (resultado.NomesCriados.Count > 0)
            {
                editor.WriteMessage("\nCatchments criados:");
                foreach (string nome in resultado.NomesCriados)
                    editor.WriteMessage($"\n  - {nome}");
            }

            if (resultado.AvisosHatchesIgnoradas.Count > 0)
            {
                editor.WriteMessage("\nHatches ignoradas:");
                foreach (string aviso in resultado.AvisosHatchesIgnoradas)
                    editor.WriteMessage($"\n  - {aviso}");
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

            // Status resumido na janela.
            string mensagemStatus;
            if (resultado.TotalCatchmentsCriados == 0)
            {
                mensagemStatus = resultado.TotalHatchesEncontradas == 0
                    ? "Nenhuma hatch encontrada na layer informada."
                    : "Nenhum catchment foi criado. Consulte a linha de comando para detalhes.";
            }
            else
            {
                mensagemStatus =
                    $"{resultado.TotalCatchmentsCriados} catchment(s) criado(s) " +
                    $"({resultado.TotalComFlowPath} com flow path).";
                if (resultado.AvisosHatchesIgnoradas.Count > 0)
                    mensagemStatus += $" {resultado.AvisosHatchesIgnoradas.Count} hatch(es) ignorada(s).";
            }

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = resultado.TotalCatchmentsCriados > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro na criacao de catchments: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    /// <summary>
    /// Lista os nomes de todas as layers do desenho atual, em ordem alfabetica.
    /// </summary>
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

    /// <summary>
    /// Lista os nomes das superficies TIN disponiveis no desenho atual.
    /// </summary>
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

    /// <summary>
    /// Lista os nomes dos Catchment Groups disponiveis no desenho atual.
    /// </summary>
    private static List<string> ObterNomesGruposCatchment(Database db)
    {
        var nomes = new List<string>();

        try
        {
            CatchmentGroupCollection grupos = CatchmentGroupCollection.GetCatchmentGroups(db);

            using Transaction transacao = db.TransactionManager.StartTransaction();

            foreach (ObjectId id in grupos)
            {
                if (id.IsNull || id.IsErased) continue;
                if (transacao.GetObject(id, OpenMode.ForRead) is CatchmentGroup grupo
                    && !string.IsNullOrWhiteSpace(grupo.Name))
                {
                    nomes.Add(grupo.Name);
                }
            }

            transacao.Commit();
        }
        catch
        {
            // Desenho sem nenhum catchment group ainda — retorna lista vazia.
        }

        return nomes.Distinct().OrderBy(x => x).ToList();
    }

    /// <summary>
    /// Lista os nomes dos Catchment Styles disponiveis no desenho atual.
    /// </summary>
    private static List<string> ObterNomesEstilosCatchment(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.CatchmentStyles)
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
            catch
            {
                // Estilo corrompido — ignora e segue.
            }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }
}
