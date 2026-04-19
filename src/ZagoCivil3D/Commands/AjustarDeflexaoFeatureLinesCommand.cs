using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
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
/// Comando que porta para C# a rotina Dynamo TER de ajuste iterativo de deflexao
/// em feature lines. Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class AjustarDeflexaoFeatureLinesCommand
{
    private static AjustarDeflexaoFeatureLinesWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_AJUSTAR_DEFLEXAO_FEATURE_LINES", CommandFlags.Session)]
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
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando ajuste de deflexao em feature lines.");

            Database banco = documento.Database;
            CivilDocument documentoCivil = CivilApplication.ActiveDocument;

            List<string> sites = ObterNomesSites(documentoCivil, banco);
            List<string> camadas = ObterNomesCamadas(banco);

            var janela = new AjustarDeflexaoFeatureLinesWindow(sites, camadas);

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

    private static void AoConfirmarJanela(object? remetente, AjustarDeflexaoFeatureLinesRequest request)
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

            AjustarDeflexaoFeatureLinesResultado resultado =
                AjustarDeflexaoFeatureLineService.Executar(documentoCivil, banco, editor, request);

            // Marcadores visuais sao opcionais: so cria se o usuario pediu e se ha algo para marcar.
            if (request.CriarMarcadoresNoDesenho && resultado.RelatorioPontos.Count > 0)
            {
                try
                {
                    MarcadoresVisuaisDeflexaoService.Criar(
                        banco,
                        editor,
                        resultado.RelatorioPontos,
                        request.NomeLayerMarcadorAjustado,
                        request.NomeLayerMarcadorFalha,
                        request.NomeLayerMarcadorInalterado,
                        request.AlturaTextoMarcador,
                        request.LimparMarcadoresAnteriores);
                }
                catch (System.Exception exMarcadores)
                {
                    editor.WriteMessage(
                        $"\n[ZagoCivil3D] Aviso: falha ao criar marcadores visuais - {exMarcadores.Message}");
                }
            }

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO AJUSTE DE DEFLEXAO =====");
            editor.WriteMessage($"\nSite: {(string.IsNullOrWhiteSpace(request.NomeSite) ? "(todos)" : request.NomeSite)}");
            editor.WriteMessage($"\nFiltro por nome: {(string.IsNullOrWhiteSpace(request.FiltroNome) ? "(nenhum)" : request.FiltroNome)}");
            editor.WriteMessage($"\nModo de otimizacao: {request.ModoOtimizacao}");
            editor.WriteMessage($"\nDeflexao limite: {request.DeflexaoLimite}");
            editor.WriteMessage($"\nPasso dinamico: min {request.PassoIncrementalAjusteCota} m / max {request.PassoMaximoDinamico} m");
            editor.WriteMessage($"\nPassadas: {request.QuantidadePassadas}, Indice maximo: {request.IndiceMaximo}");
            editor.WriteMessage($"\nFeature lines consideradas: {resultado.TotalFeatureLinesConsideradas}");
            editor.WriteMessage($"\nFeature lines ajustadas: {resultado.TotalFeatureLinesAjustadas}");
            editor.WriteMessage($"\nViolacoes iniciais: {resultado.TotalViolacoesIniciais}");
            editor.WriteMessage($"\nViolacoes finais (remanescentes): {resultado.TotalViolacoesFinais}");
            editor.WriteMessage($"\nEstacoes distintas alteradas: {resultado.TotalEstacoesAlteradas}");
            editor.WriteMessage($"\nIteracoes de ajuste aplicadas: {resultado.TotalIteracoesAplicadas}");
            editor.WriteMessage($"\n|dZ| medio aplicado: {FormatarMetros(resultado.DeltaZMedio)} / maximo: {FormatarMetros(resultado.DeltaZMaximo)}");
            editor.WriteMessage($"\nPontos ja dentro do limite: {resultado.TotalPontosDentroDoLimite}");
            editor.WriteMessage($"\nPontos nao convergidos: {resultado.TotalPontosNaoConvergidos}");

            if (resultado.NomesAjustadas.Count > 0)
            {
                editor.WriteMessage("\nFeature lines alteradas:");
                foreach (string nome in resultado.NomesAjustadas)
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

            string mensagemStatus;
            if (resultado.TotalViolacoesIniciais == 0)
            {
                mensagemStatus = "Nenhuma violacao de deflexao detectada: todas as feature lines ja atendem o criterio.";
            }
            else if (resultado.TotalViolacoesFinaisInternas == 0 && resultado.TotalViolacoesFinaisFronteira == 0)
            {
                mensagemStatus = $"Todas as {resultado.TotalViolacoesIniciais} violacao(oes) foram corrigidas. " +
                                 $"{resultado.TotalEstacoesAlteradas} estacao(oes) alterada(s).";
            }
            else
            {
                mensagemStatus =
                    $"Violacoes {resultado.TotalViolacoesIniciais} -> {resultado.TotalViolacoesFinais} " +
                    $"(internas {resultado.TotalViolacoesFinaisInternas}, fronteira {resultado.TotalViolacoesFinaisFronteira}). " +
                    $"{resultado.TotalEstacoesAlteradas} estacao(oes) alterada(s) em {resultado.TotalFeatureLinesAjustadas} FL. " +
                    $"|dZ| medio {FormatarMetros(resultado.DeltaZMedio)}, maximo {FormatarMetros(resultado.DeltaZMaximo)}. " +
                    "Veja a aba Resultados para detalhes.";
            }

            if (resultado.TotalPontosNaoConvergidos > 0)
                mensagemStatus += $" {resultado.TotalPontosNaoConvergidos} ponto(s) nao convergido(s).";

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            // Sucesso considera apenas violacoes INTERNAS: fronteiras dependem de FLs vizinhas
            // e nao sao necessariamente falhas.
            bool sucesso = resultado.MensagensErro.Count == 0
                           && resultado.TotalViolacoesFinaisInternas == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
            m_janelaAtiva?.AtualizarRelatorio(resultado);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro no ajuste de deflexao: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    /// <summary>
    /// Formata um valor em metros com 4 casas decimais usando cultura invariante.
    /// </summary>
    private static string FormatarMetros(double valor)
    {
        return valor.ToString("F4", System.Globalization.CultureInfo.InvariantCulture) + " m";
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
    /// Lista os nomes dos sites disponiveis no desenho atual (ordem alfabetica).
    /// </summary>
    private static List<string> ObterNomesSites(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId idSite in civilDoc.GetSiteIds())
        {
            if (transacao.GetObject(idSite, OpenMode.ForRead) is Site site
                && !string.IsNullOrWhiteSpace(site.Name))
            {
                nomes.Add(site.Name);
            }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }
}
