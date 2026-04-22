using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using System;
using System.Collections.Generic;
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ZagoCivil3D.Commands;

/// <summary>
/// Comando que cria regioes em todos os corredores do desenho. Porta da rotina
/// Dynamo "1.5 - CRIAR REGIOES A PARTIR DE CORREDORES". Janela modeless para
/// permitir interacao com o Civil 3D enquanto esta aberta.
/// </summary>
public class CriarRegioesCorredoresCommand
{
    private static CriarRegioesCorredoresWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_CRIAR_REGIOES_CORREDORES", CommandFlags.Session)]
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
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando criacao de regioes em corredores.");

            List<string> assemblies = CriarRegioesCorredoresService.ObterNomesAssemblies(documentoCivil, banco);

            if (assemblies.Count == 0)
            {
                editor.WriteMessage("\n[ZagoCivil3D] Nenhum assembly disponivel no desenho.");
                return;
            }

            var janela = new CriarRegioesCorredoresWindow(assemblies);
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

    private static void AoConfirmarJanela(object? remetente, CriarRegioesCorredoresRequest request)
    {
        Document? documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
        if (documento == null) return;

        Database banco = documento.Database;
        Editor editor = documento.Editor;
        CivilDocument documentoCivil = CivilApplication.ActiveDocument;

        try
        {
            using DocumentLock bloqueio = documento.LockDocument();

            CriarRegioesCorredoresResultado resultado =
                CriarRegioesCorredoresService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO CRIAR REGIOES =====");
            editor.WriteMessage($"\nAssembly aplicado: {request.NomeAssembly}");
            editor.WriteMessage($"\nCorredores encontrados: {resultado.TotalCorredores}");
            editor.WriteMessage($"\nBaselines encontradas: {resultado.TotalBaselines}");
            editor.WriteMessage($"\nBaselines com regioes criadas: {resultado.TotalBaselinesProcessadas}");
            editor.WriteMessage($"\nRegioes criadas: {resultado.TotalRegioesCriadas}");

            if (resultado.NomesRegioesCriadas.Count > 0)
            {
                editor.WriteMessage("\nRegioes:");
                foreach (string nome in resultado.NomesRegioesCriadas)
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
                foreach (string err in resultado.MensagensErro)
                    editor.WriteMessage($"\n  - {err}");
            }

            editor.WriteMessage("\n===== FIM =====");

            string status = resultado.TotalRegioesCriadas > 0
                ? $"{resultado.TotalRegioesCriadas} regiao(oes) criada(s) em {resultado.TotalBaselinesProcessadas} baseline(s)."
                : "Nenhuma regiao criada.";
            if (resultado.MensagensErro.Count > 0)
                status += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = resultado.TotalRegioesCriadas > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(status, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro durante execucao: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }
}
