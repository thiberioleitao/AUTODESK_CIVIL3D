using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ZagoCivil3D.Commands;

/// <summary>
/// Comando que cria um corredor vazio para cada alinhamento do desenho.
/// Replica a rotina Dynamo "CRIAR CORREDORES A PARTIR DE ALINHAMENTOS".
/// Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class CriarCorredoresAlinhamentosCommand
{
    private static CriarCorredoresAlinhamentosWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_CRIAR_CORREDORES_POR_ALINHAMENTOS", CommandFlags.Session)]
    public void Executar()
    {
        Document? documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
        if (documento == null)
            return;

        Database banco = documento.Database;
        Editor editor = documento.Editor;
        CivilDocument documentoCivil = CivilApplication.ActiveDocument;

        if (m_janelaAtiva != null)
        {
            m_janelaAtiva.Activate();
            return;
        }

        try
        {
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando criacao de corredores a partir de alinhamentos.");

            List<string> nomesAlinhamentos = ObterNomesAlinhamentos(documentoCivil, banco);

            var janela = new CriarCorredoresAlinhamentosWindow(nomesAlinhamentos);

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

    private static void AoConfirmarJanela(object? remetente, CriarCorredoresAlinhamentosRequest request)
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

            CriarCorredoresAlinhamentosResultado resultado =
                CorredorCreationService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO CRIAR CORREDORES =====");
            editor.WriteMessage($"\nPrefixo: '{request.PrefixoNome}'");
            editor.WriteMessage($"\nSufixo: '{request.SufixoNome}'");
            editor.WriteMessage($"\nIgnorar existentes: {(request.IgnorarExistentes ? "sim" : "nao")}");
            editor.WriteMessage($"\nAlinhamentos encontrados: {resultado.TotalAlinhamentos}");
            editor.WriteMessage($"\nCorredores criados: {resultado.TotalCorredoresCriados}");
            editor.WriteMessage($"\nIgnorados (ja existiam): {resultado.TotalIgnorados}");

            if (resultado.NomesCriados.Count > 0)
            {
                editor.WriteMessage("\nCorredores criados:");
                foreach (string nome in resultado.NomesCriados)
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

            string mensagemStatus = resultado.TotalCorredoresCriados > 0
                ? $"{resultado.TotalCorredoresCriados} corredor(es) criado(s) com sucesso."
                : "Nenhum corredor criado.";

            if (resultado.TotalIgnorados > 0)
                mensagemStatus += $" {resultado.TotalIgnorados} ignorado(s).";

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = resultado.TotalCorredoresCriados > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando de criar corredores: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    /// <summary>
    /// Le os nomes de todos os alinhamentos do desenho em ordem alfabetica.
    /// Usado apenas para preenchimento do preview na UI (primeiro nome).
    /// </summary>
    private static List<string> ObterNomesAlinhamentos(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        ObjectIdCollection idsAlinhamentos = civilDoc.GetAlignmentIds();
        if (idsAlinhamentos.Count == 0)
            return nomes;

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in idsAlinhamentos)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is Alignment alinhamento
                    && !string.IsNullOrWhiteSpace(alinhamento.Name))
                {
                    nomes.Add(alinhamento.Name);
                }
            }
            catch { }
        }

        transacao.Commit();
        return nomes.OrderBy(x => x).ToList();
    }
}
