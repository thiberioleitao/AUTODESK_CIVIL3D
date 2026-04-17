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
using ZagoCivil3D.Models;
using ZagoCivil3D.Services;
using ZagoCivil3D.Views;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ZagoCivil3D.Commands;

/// <summary>
/// Comando que cria profile views para todos os alinhamentos do desenho,
/// empilhados verticalmente a partir de uma coordenada inicial. Replica a
/// rotina Dynamo 1.3. Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class CriarProfileViewsCommand
{
    private static CriarProfileViewsWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_CRIAR_PROFILE_VIEWS", CommandFlags.Session)]
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
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando criacao de profile views.");

            List<string> estilosProfileView = ObterNomesEstilosProfileView(documentoCivil, banco);
            List<string> conjuntosBands = ObterNomesConjuntosBands(documentoCivil, banco);

            var janela = new CriarProfileViewsWindow(estilosProfileView, conjuntosBands);

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

    private static void AoConfirmarJanela(object? remetente, CriarProfileViewsRequest request)
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

            CriarProfileViewsResultado resultado =
                ProfileViewService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO PROFILE VIEWS =====");
            editor.WriteMessage($"\nAlinhamentos encontrados: {resultado.TotalAlinhamentos}");
            editor.WriteMessage($"\nProfile views criados: {resultado.TotalProfileViewsCriados}");
            editor.WriteMessage($"\nProfile views substituidos: {resultado.TotalProfileViewsSubstituidos}");
            editor.WriteMessage($"\nProfile views ignorados: {resultado.TotalProfileViewsIgnorados}");

            if (resultado.NomesCriados.Count > 0)
            {
                editor.WriteMessage("\nNomes criados:");
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

            // Exibir status na janela
            string mensagemStatus = resultado.TotalProfileViewsCriados > 0
                ? $"{resultado.TotalProfileViewsCriados} profile view(s) criado(s) com sucesso."
                : "Nenhum profile view criado.";

            if (resultado.TotalProfileViewsSubstituidos > 0)
                mensagemStatus += $" {resultado.TotalProfileViewsSubstituidos} substituido(s).";

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = resultado.TotalProfileViewsCriados > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando de profile views: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    private static List<string> ObterNomesEstilosProfileView(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.ProfileViewStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is ProfileViewStyle estilo
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

    private static List<string> ObterNomesConjuntosBands(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.ProfileViewBandSetStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is StyleBase conjunto
                    && !string.IsNullOrWhiteSpace(conjunto.Name))
                {
                    nomes.Add(conjunto.Name);
                }
            }
            catch { }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }
}
