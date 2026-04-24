using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Logica de criacao de Profile Line Labels nos profile views do desenho.
/// Replica a rotina Dynamo "ADICIONAR LABELS AS PROFILE VIEWS":
/// para cada alinhamento, lista seus profile views e profiles, filtra os
/// profiles cujo nome contem o texto informado e, para cada combinacao
/// (profile view x profile filtrado), cria um ProfileLineLabelGroup com
/// o Profile Line Label Style escolhido.
/// </summary>
public static class AdicionarLabelsProfileViewsService
{
    /// <summary>
    /// Executa o fluxo completo para todos os alinhamentos do desenho.
    /// </summary>
    public static AdicionarLabelsProfileViewsResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        AdicionarLabelsProfileViewsRequest request)
    {
        var resultado = new AdicionarLabelsProfileViewsResultado();

        ObjectId idEstilo = ObterIdProfileLineLabelStylePorNome(
            civilDoc, db, request.NomeProfileLineLabelStyle);

        if (idEstilo.IsNull)
        {
            resultado.MensagensErro.Add(
                $"Profile Line Label Style '{request.NomeProfileLineLabelStyle}' nao encontrado.");
            return resultado;
        }

        ObjectIdCollection idsAlinhamentos = civilDoc.GetAlignmentIds();
        resultado.TotalAlinhamentos = idsAlinhamentos.Count;

        if (idsAlinhamentos.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhum alinhamento encontrado no desenho.");
            return resultado;
        }

        ed.WriteMessage(
            $"\n[ZagoCivil3D] Adicionando labels '{request.NomeProfileLineLabelStyle}' com filtro '{request.FiltroNomeProfile}' em {idsAlinhamentos.Count} alinhamento(s)...");

        foreach (ObjectId idAlinhamento in idsAlinhamentos)
        {
            ProcessarAlinhamento(db, idAlinhamento, idEstilo, request, resultado, ed);
        }

        return resultado;
    }

    /// <summary>
    /// Para um alinhamento, lista profile views e profiles filtrados; cria um
    /// ProfileLineLabelGroup em cada combinacao valida. Cada alinhamento roda
    /// em sua propria transacao para isolar erros.
    /// </summary>
    private static void ProcessarAlinhamento(
        Database db,
        ObjectId idAlinhamento,
        ObjectId idEstilo,
        AdicionarLabelsProfileViewsRequest request,
        AdicionarLabelsProfileViewsResultado resultado,
        Editor ed)
    {
        string nomeAlinhamento = "?";
        string etapaAtual = "abertura";

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            etapaAtual = "leitura do alinhamento";
            if (transacao.GetObject(idAlinhamento, OpenMode.ForRead) is not Alignment alinhamento)
            {
                resultado.MensagensErro.Add("Objeto nao e um Alignment valido.");
                transacao.Commit();
                return;
            }

            nomeAlinhamento = alinhamento.Name;
            ed.WriteMessage($"\n[ZagoCivil3D] Alinhamento: {nomeAlinhamento}");

            etapaAtual = "listagem de profile views";
            List<ObjectId> idsProfileViews = ListarProfileViews(alinhamento);
            resultado.TotalProfileViews += idsProfileViews.Count;

            if (idsProfileViews.Count == 0)
            {
                resultado.Logs.Add($"'{nomeAlinhamento}': nenhum profile view.");
                transacao.Commit();
                return;
            }

            etapaAtual = "filtragem de profiles";
            List<(ObjectId id, string nome)> profilesFiltrados =
                ListarProfilesFiltrados(transacao, alinhamento, request.FiltroNomeProfile);
            resultado.TotalProfilesFiltrados += profilesFiltrados.Count;

            if (profilesFiltrados.Count == 0)
            {
                resultado.Logs.Add(
                    $"'{nomeAlinhamento}': nenhum profile com nome contendo '{request.FiltroNomeProfile}'.");
                transacao.Commit();
                return;
            }

            etapaAtual = "criacao dos label groups";
            foreach (ObjectId idProfileView in idsProfileViews)
            {
                foreach ((ObjectId idProfile, string nomeProfile) in profilesFiltrados)
                {
                    try
                    {
                        ProfileLineLabelGroup.Create(idProfileView, idProfile, idEstilo);
                        resultado.TotalLabelsCriados++;
                        resultado.Logs.Add(
                            $"'{nomeAlinhamento}': label aplicado no profile '{nomeProfile}'.");
                    }
                    catch (Exception exItem)
                    {
                        resultado.MensagensErro.Add(
                            $"'{nomeAlinhamento}' - profile '{nomeProfile}': {exItem.Message}");
                    }
                }
            }

            transacao.Commit();
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add(
                $"'{nomeAlinhamento}' - erro na etapa '{etapaAtual}': {ex.Message}");
        }
    }

    /// <summary>
    /// Retorna os ObjectIds validos dos profile views associados ao alinhamento.
    /// </summary>
    private static List<ObjectId> ListarProfileViews(Alignment alinhamento)
    {
        var ids = new List<ObjectId>();
        try
        {
            foreach (ObjectId id in alinhamento.GetProfileViewIds())
            {
                if (!id.IsNull && !id.IsErased)
                    ids.Add(id);
            }
        }
        catch
        {
            // Alinhamento sem profile views: retorna lista vazia.
        }
        return ids;
    }

    /// <summary>
    /// Retorna a lista de profiles do alinhamento cujo nome contem o filtro
    /// (case-insensitive). Quando o filtro e vazio, retorna todos os profiles.
    /// </summary>
    private static List<(ObjectId id, string nome)> ListarProfilesFiltrados(
        Transaction transacao,
        Alignment alinhamento,
        string filtro)
    {
        var profiles = new List<(ObjectId, string)>();

        ObjectIdCollection idsProfiles;
        try
        {
            idsProfiles = alinhamento.GetProfileIds();
        }
        catch
        {
            return profiles;
        }

        foreach (ObjectId idProfile in idsProfiles)
        {
            if (idProfile.IsNull || idProfile.IsErased) continue;
            try
            {
                if (transacao.GetObject(idProfile, OpenMode.ForRead) is Profile profile)
                {
                    string nome = profile.Name ?? string.Empty;
                    if (string.IsNullOrEmpty(filtro)
                        || nome.IndexOf(filtro, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        profiles.Add((idProfile, nome));
                    }
                }
            }
            catch
            {
                // Profile invalido: ignora.
            }
        }

        return profiles;
    }

    /// <summary>
    /// Localiza o Profile Line Label Style pelo nome (case-insensitive).
    /// </summary>
    private static ObjectId ObterIdProfileLineLabelStylePorNome(
        CivilDocument civilDoc,
        Database db,
        string nome)
    {
        if (string.IsNullOrWhiteSpace(nome))
            return ObjectId.Null;

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.LabelStyles.ProfileLabelStyles.LineLabelStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is StyleBase estilo
                    && string.Equals(estilo.Name, nome, StringComparison.OrdinalIgnoreCase))
                {
                    transacao.Commit();
                    return id;
                }
            }
            catch { }
        }

        transacao.Commit();
        return ObjectId.Null;
    }
}
