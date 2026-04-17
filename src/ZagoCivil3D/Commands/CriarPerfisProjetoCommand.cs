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
/// Comando que cria perfis de projeto (layout profiles) para todos os
/// alinhamentos do desenho, replicando a rotina Dynamo 1.1.
/// Usa janela modeless para nao bloquear o Civil 3D.
/// </summary>
public class CriarPerfisProjetoCommand
{
    private static CriarPerfisProjetoWindow? m_janelaAtiva;

    [CommandMethod("ZAGO_CRIAR_PERFIS_DE_PROJETO", CommandFlags.Session)]
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
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando criacao de perfis de projeto.");

            List<string> superficies = ObterNomesSuperficiesTin(documentoCivil, banco);
            List<string> estilosPerfil = ObterNomesEstilosPerfil(documentoCivil, banco);
            List<string> conjuntosRotulos = ObterNomesConjuntosRotulosPerfil(documentoCivil, banco);
            List<string> camadas = ObterTodosNomesCamadas(banco);

            var janela = new CriarPerfisProjetoWindow(superficies, estilosPerfil, conjuntosRotulos, camadas);

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

    private static void AoConfirmarJanela(object? remetente, CriarPerfisProjetoRequest request)
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

            CriarPerfisProjetoResultado resultado =
                PerfilProjetoService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO PERFIS DE PROJETO =====");
            editor.WriteMessage($"\nAlinhamentos encontrados: {resultado.TotalAlinhamentos}");
            editor.WriteMessage($"\nPerfis criados: {resultado.TotalPerfisCriados}");
            editor.WriteMessage($"\nPerfis substituidos: {resultado.TotalPerfisSubstituidos}");
            editor.WriteMessage($"\nPerfis ignorados: {resultado.TotalPerfisIgnorados}");

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
            string mensagemStatus = resultado.TotalPerfisCriados > 0
                ? $"{resultado.TotalPerfisCriados} perfil(is) criado(s) com sucesso."
                : "Nenhum perfil criado.";

            if (resultado.TotalPerfisSubstituidos > 0)
                mensagemStatus += $" {resultado.TotalPerfisSubstituidos} substituido(s).";

            if (resultado.MensagensErro.Count > 0)
                mensagemStatus += $" {resultado.MensagensErro.Count} erro(s).";

            bool sucesso = resultado.TotalPerfisCriados > 0 && resultado.MensagensErro.Count == 0;
            m_janelaAtiva?.DefinirStatus(mensagemStatus, sucesso);
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando de perfis de projeto: {ex.Message}");
            m_janelaAtiva?.DefinirStatus($"Erro: {ex.Message}", sucesso: false);
        }
    }

    private static List<string> ObterNomesSuperficiesTin(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.GetSurfaceIds())
        {
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is TinSurface superficie
                    && !string.IsNullOrWhiteSpace(superficie.Name))
                {
                    nomes.Add(superficie.Name);
                }
            }
            catch { }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }

    private static List<string> ObterNomesEstilosPerfil(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.ProfileStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is ProfileStyle estilo
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

    private static List<string> ObterNomesConjuntosRotulosPerfil(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles)
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
