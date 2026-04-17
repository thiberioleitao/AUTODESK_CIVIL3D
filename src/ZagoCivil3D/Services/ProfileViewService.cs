using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Logica de criacao automatica de profile views para todos os alinhamentos
/// do desenho, empilhados verticalmente a partir de uma coordenada inicial.
/// Replica a rotina Dynamo 1.3 "CRIAR PROFILE VIEW A PARTIR DE ALINHAMENTOS":
/// cada profile view e posicionado abaixo do anterior somando a altura real
/// (bounding box) mais um offset adicional configuravel.
/// </summary>
public static class ProfileViewService
{
    /// <summary>
    /// Executa o fluxo completo: para cada alinhamento, cria um profile view
    /// e posiciona abaixo do anterior com base na altura medida e no offset.
    /// </summary>
    public static CriarProfileViewsResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        CriarProfileViewsRequest request)
    {
        var resultado = new CriarProfileViewsResultado();

        // Estilo e conjunto de bands sao obrigatorios: a API Civil 3D
        // ProfileView.Create so oferece o overload que exige ambos.
        ObjectId idEstiloProfileView = ObterIdEstiloProfileView(civilDoc, db, request.NomeEstiloProfileView);
        if (idEstiloProfileView.IsNull)
        {
            resultado.MensagensErro.Add($"Estilo de profile view '{request.NomeEstiloProfileView}' nao encontrado.");
            return resultado;
        }

        ObjectId idConjuntoBands = ObterIdConjuntoBands(civilDoc, db, request.NomeConjuntoBands);
        if (idConjuntoBands.IsNull)
        {
            resultado.MensagensErro.Add($"Conjunto de bands '{request.NomeConjuntoBands}' nao encontrado.");
            return resultado;
        }

        ObjectIdCollection idsAlinhamentos = civilDoc.GetAlignmentIds();
        resultado.TotalAlinhamentos = idsAlinhamentos.Count;

        if (idsAlinhamentos.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhum alinhamento encontrado no desenho.");
            return resultado;
        }

        ed.WriteMessage($"\n[ZagoCivil3D] Processando {idsAlinhamentos.Count} alinhamento(s)...");

        // Posicao Y corrente: comeca na coordenada inicial e decresce conforme
        // cada profile view e criado (eixo Y do desenho, empilhamento para baixo).
        double yAtual = request.CoordenadaY;

        foreach (ObjectId idAlinhamento in idsAlinhamentos)
        {
            double? alturaCriada = ProcessarAlinhamento(
                db, idAlinhamento,
                new Point3d(request.CoordenadaX, yAtual, 0.0),
                idEstiloProfileView, idConjuntoBands,
                request, resultado, ed);

            if (alturaCriada.HasValue)
            {
                // Proximo profile view vai abaixo: desce pela altura real
                // medida no extents + o offset adicional definido pelo usuario.
                yAtual -= alturaCriada.Value + request.OffsetAdicional;
            }
        }

        return resultado;
    }

    /// <summary>
    /// Processa um unico alinhamento: cria o profile view no ponto informado,
    /// mede sua altura e retorna o valor. Retorna null se o alinhamento foi
    /// ignorado ou houve erro.
    /// </summary>
    private static double? ProcessarAlinhamento(
        Database db,
        ObjectId idAlinhamento,
        Point3d pontoInsercao,
        ObjectId idEstiloProfileView,
        ObjectId idConjuntoBands,
        CriarProfileViewsRequest request,
        CriarProfileViewsResultado resultado,
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
                return null;
            }

            nomeAlinhamento = alinhamento.Name;
            string nomeProfileView = nomeAlinhamento + request.SufixoProfileView;

            ed.WriteMessage($"\n[ZagoCivil3D] Alinhamento: {nomeAlinhamento}");

            // Verificar se profile view ja existe
            etapaAtual = "verificacao de profile view existente";
            ObjectId idPvExistente = ObterIdProfileViewPorNome(transacao, alinhamento, nomeProfileView);

            if (!idPvExistente.IsNull)
            {
                if (!request.SubstituirExistentes)
                {
                    resultado.TotalProfileViewsIgnorados++;
                    resultado.Logs.Add($"'{nomeAlinhamento}': profile view '{nomeProfileView}' ja existe, ignorado.");
                    transacao.Commit();
                    return null;
                }

                etapaAtual = "exclusao do profile view existente";
                ProfileView pvAntigo = (ProfileView)transacao.GetObject(idPvExistente, OpenMode.ForWrite);
                pvAntigo.Erase();
                resultado.TotalProfileViewsSubstituidos++;
                resultado.Logs.Add($"'{nomeAlinhamento}': profile view '{nomeProfileView}' existente removido.");
            }

            // Criar profile view com estilo e conjunto de bands especificados.
            etapaAtual = "criacao do profile view";
            ObjectId idProfileView = ProfileView.Create(
                alinhamento.ObjectId,
                pontoInsercao,
                nomeProfileView,
                idConjuntoBands,
                idEstiloProfileView);

            etapaAtual = "leitura do profile view para medir altura";
            ProfileView profileView = (ProfileView)transacao.GetObject(idProfileView, OpenMode.ForRead);

            // Garantir faixas automaticas de estacao e elevacao, como na rotina
            // Dynamo (SetElevationRangeAutomatic(true), SetStationRangeAutomatic(true)).
            etapaAtual = "configuracao de faixas automaticas";
            profileView.UpgradeOpen();
            try
            {
                profileView.ElevationRangeMode = ElevationRangeType.Automatic;
                profileView.StationRangeMode = StationRangeType.Automatic;
            }
            catch
            {
                // Nao fatal: se a propriedade nao estiver disponivel, segue em frente.
            }
            profileView.DowngradeOpen();

            double altura = MedirAlturaProfileView(profileView);

            transacao.Commit();

            resultado.TotalProfileViewsCriados++;
            resultado.NomesCriados.Add(nomeProfileView);
            ed.WriteMessage($"\n[ZagoCivil3D]   Profile view '{nomeProfileView}' criado (altura={altura:F2}).");

            return altura;
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add($"'{nomeAlinhamento}' - erro na etapa '{etapaAtual}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Retorna a altura (bounding box em Y) do profile view, em unidades do
    /// desenho. Retorna 0 se o GeometricExtents nao estiver disponivel.
    /// </summary>
    private static double MedirAlturaProfileView(ProfileView profileView)
    {
        try
        {
            Extents3d extensao = profileView.GeometricExtents;
            double altura = extensao.MaxPoint.Y - extensao.MinPoint.Y;
            return altura > 0 ? altura : 0.0;
        }
        catch
        {
            return 0.0;
        }
    }

    /// <summary>
    /// Retorna o ObjectId do profile view do alinhamento com o nome informado,
    /// ou ObjectId.Null se nao existir.
    /// </summary>
    private static ObjectId ObterIdProfileViewPorNome(
        Transaction transacao,
        Alignment alinhamento,
        string nomeProfileView)
    {
        foreach (ObjectId idPv in alinhamento.GetProfileViewIds())
        {
            try
            {
                if (transacao.GetObject(idPv, OpenMode.ForRead) is ProfileView pv
                    && string.Equals(pv.Name, nomeProfileView, StringComparison.OrdinalIgnoreCase))
                {
                    return idPv;
                }
            }
            catch
            {
                // Profile view invalido/apagado, ignorar
            }
        }

        return ObjectId.Null;
    }

    /// <summary>
    /// Localiza o estilo de profile view pelo nome.
    /// </summary>
    private static ObjectId ObterIdEstiloProfileView(CivilDocument civilDoc, Database db, string nome)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.ProfileViewStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is ProfileViewStyle estilo
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

    /// <summary>
    /// Localiza o conjunto de bands de profile view pelo nome.
    /// </summary>
    private static ObjectId ObterIdConjuntoBands(CivilDocument civilDoc, Database db, string nome)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.ProfileViewBandSetStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is StyleBase conjunto
                    && string.Equals(conjunto.Name, nome, StringComparison.OrdinalIgnoreCase))
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
