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

namespace ZagoCivil3D.Commands;

/// <summary>
/// Comando que porta para C# o fluxo do Dynamo de ajuste de terraplenagem.
/// </summary>
public class TerraplenagemFeatureLinesCommand
{
    /// <summary>
    /// Comando exposto ao Civil 3D para abrir a UI e executar o processamento.
    /// </summary>
    [CommandMethod("ZAGO_TERRAPLENAGEM_FEATURE_LINES_SEPARADAS")]
    public void Executar()
    {
        Document? documento = Application.DocumentManager.MdiActiveDocument;
        if (documento == null)
            return;

        Database banco = documento.Database;
        Editor editor = documento.Editor;
        CivilDocument documentoCivil = CivilApplication.ActiveDocument;

        try
        {
            editor.WriteMessage("\n[ZagoCivil3D] Iniciando fluxo de terraplenagem das feature lines.");

            TerraplenagemFeatureLinesRequest? request = ColetarParametros(documentoCivil, banco);
            if (request == null)
            {
                editor.WriteMessage("\n[ZagoCivil3D] Comando cancelado pelo usuario.");
                return;
            }

            TerraplenagemFeatureLinesResultado resultado =
                TerraplenagemFeatureLineService.Executar(documentoCivil, banco, editor, request);

            editor.WriteMessage("\n");
            editor.WriteMessage("\n===== RESUMO TERRAPLENAGEM =====");
            editor.WriteMessage($"\nFeature lines no site: {resultado.TotalFeatureLinesNoSite}");
            editor.WriteMessage($"\nFeature lines filtradas pelo poligono: {resultado.TotalFeatureLinesFiltradas}");
            editor.WriteMessage($"\nFeature lines com ajuste de deflexao: {resultado.TotalFeatureLinesComDeflexaoAjustada}");
            editor.WriteMessage($"\nFeature lines com ajuste por superficie: {resultado.TotalFeatureLinesComSuperficieAjustada}");

            if (resultado.Logs.Count > 0)
            {
                editor.WriteMessage("\nLogs:");
                foreach (string log in resultado.Logs)
                    editor.WriteMessage($"\n - {log}");
            }

            if (resultado.MensagensErro.Count > 0)
            {
                editor.WriteMessage("\nMensagens/erros:");
                foreach (string mensagem in resultado.MensagensErro)
                    editor.WriteMessage($"\n - {mensagem}");
            }

            editor.WriteMessage("\n===== FIM =====");
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\n[ZagoCivil3D] Erro no comando de terraplenagem: {ex.Message}");
        }
    }

    /// <summary>
    /// Monta a UI usando as listas atuais do desenho e devolve os parametros informados.
    /// </summary>
    private static TerraplenagemFeatureLinesRequest? ColetarParametros(CivilDocument civilDoc, Database db)
    {
        List<string> sites = ObterNomesSites(civilDoc, db);
        List<string> camadas = ObterTodosNomesCamadas(db);
        List<string> superficies = ObterNomesSuperficiesTin(civilDoc, db);

        var tela = new TerraplenagemFeatureLinesWindow(sites, camadas, superficies);
        bool? resultadoDialogo = tela.ShowDialog();

        if (resultadoDialogo != true)
            return null;

        return tela.CriarRequest();
    }

    /// <summary>
    /// Lista os nomes de todas as layers do desenho atual.
    /// </summary>
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

    /// <summary>
    /// Lista os nomes dos sites disponiveis no desenho atual.
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

    /// <summary>
    /// Lista apenas superficies TIN, porque essa e a mesma restricao usada no Dynamo.
    /// </summary>
    private static List<string> ObterNomesSuperficiesTin(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId idSuperficie in civilDoc.GetSurfaceIds())
        {
            if (transacao.GetObject(idSuperficie, OpenMode.ForRead) is TinSurface superficie
                && !string.IsNullOrWhiteSpace(superficie.Name))
            {
                nomes.Add(superficie.Name);
            }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }
}
