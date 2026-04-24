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
/// Comando que converte todos os alinhamentos do Civil 3D em polilinhas 2D
/// do AutoCAD. Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class ConverterAlinhamentosEmPolilinhasCommand
{
    private static ConverterAlinhamentosEmPolilinhasWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_CONVERTER_ALINHAMENTOS_EM_POLILINHAS", CommandFlags.Session)]
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
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando conversao de alinhamentos em polilinhas 2D.");

            Database banco = documento.Database;
            List<string> camadas = ObterNomesCamadas(banco);

            var janela = new ConverterAlinhamentosEmPolilinhasWindow(camadas);

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

    private static void AoConfirmarJanela(
        object? remetente,
        ConverterAlinhamentosEmPolilinhasRequest request)
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

            ConverterAlinhamentosEmPolilinhasResultado resultado =
                ConverterAlinhamentosEmPolilinhasService.Executar(
                    documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO CONVERTER ALINHAMENTOS EM POLILINHAS =====");
            editor.WriteMessage($"\nFiltro de nome: {(string.IsNullOrWhiteSpace(request.FiltroNome) ? "(nenhum)" : $"contem '{request.FiltroNome}'")}");
            editor.WriteMessage($"\nLayer destino: {(string.IsNullOrWhiteSpace(request.NomeLayer) ? "(corrente)" : request.NomeLayer)}");
            editor.WriteMessage($"\nElevation: {request.Elevation}");
            editor.WriteMessage($"\nPreservar arcos: {(request.PreservarArcos ? "sim" : "nao")}");
            editor.WriteMessage($"\nPasso discretizacao: {request.PassoDiscretizacaoEspirais} m");
            editor.WriteMessage($"\nApagar alinhamentos: {(request.ApagarAlinhamentosOriginais ? "sim" : "nao")}");
            editor.WriteMessage($"\nDry-run: {(request.DryRun ? "sim" : "nao")}");
            editor.WriteMessage($"\nAlinhamentos encontrados: {resultado.TotalAlinhamentos}");
            editor.WriteMessage($"\nAlinhamentos apos filtro: {resultado.TotalAlinhamentosFiltrados}");
            editor.WriteMessage($"\nPolilinhas criadas: {resultado.TotalPolilinhasCriadas}");
            editor.WriteMessage($"\nAlinhamentos apagados: {resultado.TotalAlinhamentosApagados}");

            if (resultado.Convertidos.Count > 0)
            {
                editor.WriteMessage("\nConvertidos:");
                foreach (AlinhamentoConvertido c in resultado.Convertidos)
                {
                    editor.WriteMessage(
                        $"\n  - {c.NomeAlinhamento}: {c.TotalVertices} vertice(s). " +
                        $"Linhas={c.TotalLinhas} Arcos={c.TotalArcos} Espirais={c.TotalEspirais} " +
                        $"(+{c.VerticesDiscretizadosEspirais} vert. em espirais). " +
                        $"L={c.ComprimentoAlinhamento:F2} m.");
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
                           && (request.DryRun
                               ? resultado.Convertidos.Count > 0
                               : resultado.TotalPolilinhasCriadas > 0);

            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro na conversao: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    private static string MontarMensagemStatus(
        ConverterAlinhamentosEmPolilinhasRequest request,
        ConverterAlinhamentosEmPolilinhasResultado resultado)
    {
        if (resultado.TotalAlinhamentos == 0)
            return "Nenhum alinhamento encontrado no desenho.";

        bool temFiltro = !string.IsNullOrWhiteSpace(request.FiltroNome);
        string origem = temFiltro
            ? $"{resultado.TotalAlinhamentosFiltrados} de {resultado.TotalAlinhamentos} alinhamento(s) (filtro '{request.FiltroNome}')"
            : $"{resultado.TotalAlinhamentos} alinhamento(s)";

        string baseTexto = request.DryRun
            ? $"{resultado.Convertidos.Count} alinhamento(s) avaliado(s) em dry-run ({origem})."
            : $"{resultado.TotalPolilinhasCriadas} polilinha(s) criada(s) a partir de {origem}.";

        if (request.ApagarAlinhamentosOriginais && !request.DryRun)
            baseTexto += $" {resultado.TotalAlinhamentosApagados} alinhamento(s) apagado(s).";

        if (resultado.Avisos.Count > 0)
            baseTexto += $" {resultado.Avisos.Count} aviso(s).";

        if (resultado.MensagensErro.Count > 0)
            baseTexto += $" {resultado.MensagensErro.Count} erro(s).";

        return baseTexto;
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
