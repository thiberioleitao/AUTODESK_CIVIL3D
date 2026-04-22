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
/// Comando que porta para C# a rotina Dynamo "CRIAR PONTOS NOS CRUZAMENTOS
/// ENTRE ALINHAMENTOS". Calcula as interseccoes entre todos os pares de
/// alinhamentos, agrupa em blocos (trackers) e cria CogoPoints com rotulos
/// sequenciais. Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class CriarPontosCruzamentosCommand
{
    private static CriarPontosCruzamentosWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_CRIAR_PONTOS_CRUZAMENTOS_ALINHAMENTOS", CommandFlags.Session)]
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
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando criacao de pontos nos cruzamentos de alinhamentos.");

            Database banco = documento.Database;
            List<string> camadas = ObterNomesCamadas(banco);

            var janela = new CriarPontosCruzamentosWindow(camadas);

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

    private static void AoConfirmarJanela(object? remetente, CriarPontosCruzamentosRequest request)
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

            CriarPontosCruzamentosResultado resultado =
                CriarPontosCruzamentosService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO CRIAR PONTOS NOS CRUZAMENTOS =====");
            editor.WriteMessage($"\nComprimento tracker: {request.ComprimentoTracker}");
            editor.WriteMessage($"\nPontos por coluna: {request.PontosPorTracker}");
            editor.WriteMessage($"\nPrefixo: {request.Prefixo}");
            editor.WriteMessage($"\nAlternar lado inicial: {(request.AlternarLadoInicial ? "sim" : "nao")}");
            editor.WriteMessage($"\nLayer: {(string.IsNullOrWhiteSpace(request.NomeLayer) ? "(corrente)" : request.NomeLayer)}");
            editor.WriteMessage($"\nCriar CogoPoints: {(request.CriarCogoPoints ? "sim" : "nao (dry-run)")}");
            editor.WriteMessage($"\nAlinhamentos: {resultado.TotalAlinhamentos}");
            editor.WriteMessage($"\nPares testados: {resultado.TotalParesTestados}");
            editor.WriteMessage($"\nPares com cruzamento: {resultado.TotalParesComIntersecao}");
            editor.WriteMessage($"\nPontos brutos: {resultado.TotalPontosBrutos}");
            editor.WriteMessage($"\nPontos unicos: {resultado.TotalPontosUnicos}");
            editor.WriteMessage($"\nPontos ordenados: {resultado.TotalPontosOrdenados}");
            editor.WriteMessage($"\nBlocos gerados: {resultado.Blocos.Count}");
            editor.WriteMessage($"\nCogoPoints criados: {resultado.TotalCogoPointsCriados}");
            editor.WriteMessage($"\nRawDescriptions aplicadas: {resultado.TotalRawDescriptionAplicadas}");

            if (resultado.Blocos.Count > 0)
            {
                editor.WriteMessage("\nBlocos:");
                foreach (BlocoCruzamento bloco in resultado.Blocos)
                {
                    editor.WriteMessage(
                        $"\n  - Bloco {bloco.IndiceBloco} ({bloco.LadoInicial}): {bloco.QuantidadePontos} ponto(s). " +
                        $"Base X={bloco.PontoBaseX:F3} Y={bloco.PontoBaseY:F3}. " +
                        $"Rotulos: {PrimeiroEUltimo(bloco.Rotulos)}.");
                }
            }

            if (resultado.Avisos.Count > 0)
            {
                editor.WriteMessage("\nAvisos:");
                foreach (string aviso in resultado.Avisos)
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

            string mensagemStatus = MontarMensagemStatus(request, resultado);
            bool sucesso = resultado.MensagensErro.Count == 0
                           && (request.CriarCogoPoints
                               ? resultado.TotalCogoPointsCriados > 0
                               : resultado.TotalPontosUnicos > 0);

            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro na criacao de pontos: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    private static string MontarMensagemStatus(
        CriarPontosCruzamentosRequest request,
        CriarPontosCruzamentosResultado resultado)
    {
        if (resultado.TotalPontosUnicos == 0)
            return "Nenhum cruzamento encontrado entre os alinhamentos.";

        string baseTexto = request.CriarCogoPoints
            ? $"{resultado.TotalCogoPointsCriados} CogoPoint(s) criado(s) em {resultado.Blocos.Count} bloco(s)."
            : $"{resultado.TotalPontosUnicos} cruzamento(s) em {resultado.Blocos.Count} bloco(s) (dry-run).";

        if (resultado.Avisos.Count > 0)
            baseTexto += $" {resultado.Avisos.Count} aviso(s).";

        if (resultado.MensagensErro.Count > 0)
            baseTexto += $" {resultado.MensagensErro.Count} erro(s).";

        return baseTexto;
    }

    private static string PrimeiroEUltimo(IReadOnlyList<string> rotulos)
    {
        if (rotulos.Count == 0) return "(vazio)";
        if (rotulos.Count == 1) return rotulos[0];
        return $"{rotulos[0]} ... {rotulos[^1]}";
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
}
