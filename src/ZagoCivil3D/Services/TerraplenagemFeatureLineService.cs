using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

/// <summary>
/// Centraliza a portabilidade do fluxo Dynamo para uma implementacao nativa em C#.
/// </summary>
public static class TerraplenagemFeatureLineService
{
    /// <summary>
    /// Tolerancia XY usada para mapear um PI point na colecao AllPoints.
    /// O valor e pequeno, mas tolera diferencas numericas residuais.
    /// </summary>
    private const double ToleranciaMapeamentoXY = 1e-4;

    /// <summary>
    /// Tolerancia usada para considerar um ponto sobre a borda do poligono.
    /// </summary>
    private const double ToleranciaBordaPoligono = 1e-8;

    /// <summary>
    /// Executa o fluxo completo de terraplenagem.
    /// A leitura inicial e separada da escrita para evitar uma transacao unica muito longa.
    /// </summary>
    public static TerraplenagemFeatureLinesResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        TerraplenagemFeatureLinesRequest request)
    {
        var resultado = new TerraplenagemFeatureLinesResultado();

        PreparacaoTerraplenagem? preparacao = PrepararExecucao(civilDoc, db, request, resultado);
        if (preparacao == null)
            return resultado;

        // Esses totais sao exibidos ao usuario ao final do comando.
        resultado.TotalFeatureLinesNoSite = preparacao.TotalFeatureLinesNoSite;
        resultado.TotalFeatureLinesFiltradas = preparacao.IdsFeatureLinesFiltradas.Count;

        if (preparacao.IdsFeatureLinesFiltradas.Count == 0)
        {
            resultado.MensagensErro.Add(
                "Nenhuma feature line do site permaneceu totalmente dentro do poligono informado.");
            return resultado;
        }

        foreach (ObjectId idFeatureLine in preparacao.IdsFeatureLinesFiltradas)
            ProcessarFeatureLine(db, preparacao.IdSuperficieBase, idFeatureLine, request, resultado);

        ed.WriteMessage("\n[ZagoCivil3D] Fluxo de terraplenagem concluido.");
        return resultado;
    }

    /// <summary>
    /// Reune todos os dados de leitura necessarios antes de iniciar as alteracoes.
    /// </summary>
    private static PreparacaoTerraplenagem? PrepararExecucao(
        CivilDocument civilDoc,
        Database db,
        TerraplenagemFeatureLinesRequest request,
        TerraplenagemFeatureLinesResultado resultado)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        Site? site = ObterSitePorNome(civilDoc, transacao, request.NomeSite);
        if (site == null)
        {
            resultado.MensagensErro.Add($"Site '{request.NomeSite}' nao encontrado.");
            return null;
        }

        TinSurface? superficieBase = ObterSuperficiePorNome(civilDoc, transacao, request.NomeSuperficieBase);
        if (superficieBase == null)
        {
            resultado.MensagensErro.Add($"Superficie base '{request.NomeSuperficieBase}' nao encontrada.");
            return null;
        }

        ResultadoSelecaoPoligono selecaoPoligono =
            SelecionarPoligonoDaLayer(db, transacao, request.NomeCamadaPoligono);

        if (!selecaoPoligono.Sucesso)
        {
            resultado.MensagensErro.Add(selecaoPoligono.Mensagem);
            return null;
        }

        List<ObjectId> idsFeatureLines = site.GetFeatureLineIds().Cast<ObjectId>().ToList();
        List<ObjectId> idsFiltrados = FiltrarFeatureLinesDentroPoligono(
            transacao,
            idsFeatureLines,
            selecaoPoligono.Vertices,
            resultado.Logs);

        transacao.Commit();

        return new PreparacaoTerraplenagem(
            superficieBase.ObjectId,
            idsFeatureLines.Count,
            idsFiltrados);
    }

    /// <summary>
    /// Processa uma unica feature line em sua propria transacao de escrita.
    /// </summary>
    private static void ProcessarFeatureLine(
        Database db,
        ObjectId idSuperficieBase,
        ObjectId idFeatureLine,
        TerraplenagemFeatureLinesRequest request,
        TerraplenagemFeatureLinesResultado resultado)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        if (transacao.GetObject(idFeatureLine, OpenMode.ForWrite) is not FeatureLine featureLine)
        {
            resultado.MensagensErro.Add(
                $"Nao foi possivel abrir a feature line {idFeatureLine.Handle} para escrita.");
            return;
        }

        if (transacao.GetObject(idSuperficieBase, OpenMode.ForRead) is not TinSurface superficieBase)
        {
            resultado.MensagensErro.Add("Nao foi possivel reabrir a superficie base para leitura.");
            return;
        }

        try
        {
            bool ajustouDeflexao = AjustarDeflexao(featureLine, request, resultado.Logs);
            if (ajustouDeflexao)
                resultado.TotalFeatureLinesComDeflexaoAjustada++;

            bool ajustouSuperficie = AjustarPorSuperficie(featureLine, superficieBase, request, resultado.Logs);
            if (ajustouSuperficie)
                resultado.TotalFeatureLinesComSuperficieAjustada++;

            transacao.Commit();
        }
        catch (System.Exception ex)
        {
            resultado.MensagensErro.Add(
                $"Erro ao processar a feature line '{featureLine.Name}': {ex.Message}");
        }
    }

    /// <summary>
    /// Recupera o site pelo nome exibido ao usuario.
    /// </summary>
    private static Site? ObterSitePorNome(CivilDocument civilDoc, Transaction transacao, string nomeSite)
    {
        foreach (ObjectId idSite in civilDoc.GetSiteIds())
        {
            if (transacao.GetObject(idSite, OpenMode.ForRead) is not Site site)
                continue;

            if (string.Equals(site.Name, nomeSite, StringComparison.OrdinalIgnoreCase))
                return site;
        }

        return null;
    }

    /// <summary>
    /// Recupera a superficie TIN selecionada para servir de referencia na terraplenagem.
    /// </summary>
    private static TinSurface? ObterSuperficiePorNome(CivilDocument civilDoc, Transaction transacao, string nomeSuperficie)
    {
        foreach (ObjectId idSuperficie in civilDoc.GetSurfaceIds())
        {
            if (transacao.GetObject(idSuperficie, OpenMode.ForRead) is not TinSurface superficie)
                continue;

            if (string.Equals(superficie.Name, nomeSuperficie, StringComparison.OrdinalIgnoreCase))
                return superficie;
        }

        return null;
    }

    /// <summary>
    /// Seleciona o poligono fechado da layer informada.
    /// O metodo falha explicitamente quando nao existe poligono valido
    /// ou quando existe mais de um candidato valido na mesma layer.
    /// </summary>
    private static ResultadoSelecaoPoligono SelecionarPoligonoDaLayer(
        Database db,
        Transaction transacao,
        string nomeCamada)
    {
        List<CandidatoPoligono> candidatos = ObterPoligonosValidosDaLayer(db, transacao, nomeCamada);

        if (candidatos.Count == 0)
        {
            return ResultadoSelecaoPoligono.Falha(
                $"Nenhum poligono fechado valido foi encontrado na layer '{nomeCamada}'.");
        }

        if (candidatos.Count > 1)
        {
            string handles = string.Join(", ", candidatos.Select(x => $"{x.Tipo} {x.Handle}"));
            return ResultadoSelecaoPoligono.Falha(
                $"A layer '{nomeCamada}' possui mais de um poligono fechado valido ({handles}). Ajuste a layer ou mantenha apenas um poligono de filtro.");
        }

        return ResultadoSelecaoPoligono.SucessoCom(candidatos[0].Vertices);
    }

    /// <summary>
    /// Localiza todas as entidades fechadas suportadas que podem servir como poligono de filtro.
    /// </summary>
    private static List<CandidatoPoligono> ObterPoligonosValidosDaLayer(
        Database db,
        Transaction transacao,
        string nomeCamada)
    {
        var candidatos = new List<CandidatoPoligono>();
        var tabelaBlocos = (BlockTable)transacao.GetObject(db.BlockTableId, OpenMode.ForRead);
        var espacoModelo =
            (BlockTableRecord)transacao.GetObject(tabelaBlocos[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        foreach (ObjectId idEntidade in espacoModelo)
        {
            if (transacao.GetObject(idEntidade, OpenMode.ForRead) is not Autodesk.AutoCAD.DatabaseServices.Entity entidade)
                continue;

            if (!string.Equals(entidade.Layer, nomeCamada, StringComparison.OrdinalIgnoreCase))
                continue;

            List<Point2d> vertices = entidade switch
            {
                Polyline polyline => ExtrairVerticesPolyline(polyline),
                Polyline2d polyline2d => ExtrairVerticesPolyline2d(polyline2d, transacao),
                Polyline3d polyline3d => ExtrairVerticesPolyline3d(polyline3d, transacao),
                _ => new List<Point2d>()
            };

            if (vertices.Count < 3)
                continue;

            candidatos.Add(
                new CandidatoPoligono(
                    entidade.GetType().Name,
                    entidade.Handle.ToString(),
                    vertices));
        }

        return candidatos;
    }

    /// <summary>
    /// Extrai os vertices de uma polyline 2D fechada.
    /// </summary>
    private static List<Point2d> ExtrairVerticesPolyline(Polyline polyline)
    {
        var vertices = new List<Point2d>();

        if (!polyline.Closed || polyline.NumberOfVertices < 3)
            return vertices;

        for (int indice = 0; indice < polyline.NumberOfVertices; indice++)
            vertices.Add(polyline.GetPoint2dAt(indice));

        return vertices;
    }

    /// <summary>
    /// Extrai os vertices de uma polyline2d fechada.
    /// </summary>
    private static List<Point2d> ExtrairVerticesPolyline2d(Polyline2d polyline, Transaction transacao)
    {
        var vertices = new List<Point2d>();

        if (!polyline.Closed)
            return vertices;

        foreach (ObjectId idVertice in polyline)
        {
            if (transacao.GetObject(idVertice, OpenMode.ForRead) is Vertex2d vertice)
                vertices.Add(new Point2d(vertice.Position.X, vertice.Position.Y));
        }

        return vertices;
    }

    /// <summary>
    /// Extrai os vertices de uma polyline3d fechada convertendo os pontos para XY.
    /// </summary>
    private static List<Point2d> ExtrairVerticesPolyline3d(Polyline3d polyline, Transaction transacao)
    {
        var vertices = new List<Point2d>();

        if (!polyline.Closed)
            return vertices;

        foreach (ObjectId idVertice in polyline)
        {
            if (transacao.GetObject(idVertice, OpenMode.ForRead) is PolylineVertex3d vertice)
                vertices.Add(new Point2d(vertice.Position.X, vertice.Position.Y));
        }

        return vertices;
    }

    /// <summary>
    /// Mantem apenas as feature lines cujos pontos relevantes permanecem dentro do poligono.
    /// O filtro usa todos os AllPoints quando disponiveis e cai para PI points apenas como fallback.
    /// </summary>
    private static List<ObjectId> FiltrarFeatureLinesDentroPoligono(
        Transaction transacao,
        IReadOnlyCollection<ObjectId> idsFeatureLines,
        IReadOnlyList<Point2d> verticesPoligono,
        ICollection<string> logs)
    {
        var filtradas = new List<ObjectId>();

        foreach (ObjectId idFeatureLine in idsFeatureLines)
        {
            if (transacao.GetObject(idFeatureLine, OpenMode.ForRead) is not FeatureLine featureLine)
                continue;

            Point3dCollection pontosRelevantes = ObterPontosParaFiltroEspacial(featureLine);
            if (pontosRelevantes.Count == 0)
            {
                logs.Add(
                    $"'{featureLine.Name}': falha geometrica no filtro espacial, a feature line nao possui pontos analisaveis.");
                continue;
            }

            bool todosDentro = true;
            foreach (Point3d ponto in pontosRelevantes)
            {
                if (PontoDentroPoligono(new Point2d(ponto.X, ponto.Y), verticesPoligono))
                    continue;

                todosDentro = false;
                break;
            }

            if (todosDentro)
                filtradas.Add(idFeatureLine);
        }

        return filtradas;
    }

    /// <summary>
    /// Retorna a colecao de pontos usada no filtro espacial.
    /// </summary>
    private static Point3dCollection ObterPontosParaFiltroEspacial(FeatureLine featureLine)
    {
        Point3dCollection todosOsPontos = featureLine.GetPoints(FeatureLinePointType.AllPoints);
        if (todosOsPontos.Count > 0)
            return todosOsPontos;

        return featureLine.GetPoints(FeatureLinePointType.PIPoint);
    }

    /// <summary>
    /// Implementa o algoritmo de ray casting usado no script Python do Dynamo.
    /// Pontos sobre a borda do poligono sao tratados como internos.
    /// </summary>
    private static bool PontoDentroPoligono(Point2d ponto, IReadOnlyList<Point2d> poligono)
    {
        for (int indice = 0; indice < poligono.Count; indice++)
        {
            Point2d inicio = poligono[indice];
            Point2d fim = poligono[(indice + 1) % poligono.Count];

            if (PontoSobreSegmento(ponto, inicio, fim))
                return true;
        }

        bool dentro = false;
        int j = poligono.Count - 1;

        for (int i = 0; i < poligono.Count; i++)
        {
            Point2d pi = poligono[i];
            Point2d pj = poligono[j];

            bool intersecta =
                ((pi.Y > ponto.Y) != (pj.Y > ponto.Y))
                && (ponto.X < (pj.X - pi.X) * (ponto.Y - pi.Y) / ((pj.Y - pi.Y) + 1e-12) + pi.X);

            if (intersecta)
                dentro = !dentro;

            j = i;
        }

        return dentro;
    }

    /// <summary>
    /// Verifica se um ponto esta sobre uma aresta do poligono.
    /// </summary>
    private static bool PontoSobreSegmento(Point2d ponto, Point2d inicio, Point2d fim)
    {
        double areaDobrada =
            ((fim.X - inicio.X) * (ponto.Y - inicio.Y))
            - ((fim.Y - inicio.Y) * (ponto.X - inicio.X));

        if (Math.Abs(areaDobrada) > ToleranciaBordaPoligono)
            return false;

        double minX = Math.Min(inicio.X, fim.X) - ToleranciaBordaPoligono;
        double maxX = Math.Max(inicio.X, fim.X) + ToleranciaBordaPoligono;
        double minY = Math.Min(inicio.Y, fim.Y) - ToleranciaBordaPoligono;
        double maxY = Math.Max(inicio.Y, fim.Y) + ToleranciaBordaPoligono;

        return ponto.X >= minX
            && ponto.X <= maxX
            && ponto.Y >= minY
            && ponto.Y <= maxY;
    }

    /// <summary>
    /// Reproduz o ajuste iterativo de deflexao do primeiro script Python.
    /// </summary>
    private static bool AjustarDeflexao(
        FeatureLine featureLine,
        TerraplenagemFeatureLinesRequest request,
        ICollection<string> logs)
    {
        bool houveAlteracao = false;

        for (int passadaGlobal = 1; passadaGlobal <= request.QuantidadePassadasGlobais; passadaGlobal++)
        {
            Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
            int quantidadePontosPi = pontosPi.Count;

            if (quantidadePontosPi < 3)
            {
                logs.Add($"'{featureLine.Name}': falha geometrica na etapa de deflexao, a feature line possui menos de 3 PI points.");
                break;
            }

            int ultimoIndiceAvaliavel = Math.Min(
                request.NumeroMaximoPontosPorFeatureLine,
                quantidadePontosPi - 2);

            if (ultimoIndiceAvaliavel < 1)
                break;

            IEnumerable<int> indices = passadaGlobal % 2 == 1
                ? Enumerable.Range(1, ultimoIndiceAvaliavel)
                : Enumerable.Range(1, ultimoIndiceAvaliavel).Reverse();

            int pontosOk = 0;
            int totalPontos = ultimoIndiceAvaliavel;

            foreach (int indicePi in indices)
            {
                ResultadoEtapa resultadoPonto = ProcessarPontoDeflexao(
                    featureLine,
                    indicePi,
                    request,
                    ref houveAlteracao);

                if (resultadoPonto.Tipo == TipoResultadoEtapa.AjusteConcluido)
                {
                    pontosOk++;
                    continue;
                }

                logs.Add($"'{featureLine.Name}': PI {indicePi} - {resultadoPonto.Mensagem}");
            }

            double percentualOk = totalPontos == 0 ? 1 : (double)pontosOk / totalPontos;
            logs.Add(
                $"'{featureLine.Name}': passada {passadaGlobal}, pontos dentro do limite = {pontosOk}/{totalPontos}, percentual = {percentualOk.ToString("P2", CultureInfo.InvariantCulture)}.");

            if (percentualOk >= request.PercentualObjetivo)
                break;
        }

        return houveAlteracao;
    }

    /// <summary>
    /// Processa um unico PI point na etapa de deflexao.
    /// </summary>
    private static ResultadoEtapa ProcessarPontoDeflexao(
        FeatureLine featureLine,
        int indicePi,
        TerraplenagemFeatureLinesRequest request,
        ref bool houveAlteracao)
    {
        for (int tentativaLocal = 1; tentativaLocal <= request.NumeroTentativasPorPonto; tentativaLocal++)
        {
            Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);

            Point3d pontoAnterior = pontosPi[indicePi - 1];
            Point3d pontoAtual = pontosPi[indicePi];
            Point3d pontoPosterior = pontosPi[indicePi + 1];

            double comprimentoAnterior = Distancia2d(pontoAnterior, pontoAtual);
            double comprimentoPosterior = Distancia2d(pontoAtual, pontoPosterior);

            if (comprimentoAnterior <= ToleranciaMapeamentoXY || comprimentoPosterior <= ToleranciaMapeamentoXY)
            {
                return ResultadoEtapa.FalhaGeometrica(
                    "falha geometrica na etapa de deflexao: um dos segmentos adjacentes possui comprimento 2D nulo.");
            }

            double deflexao =
                (pontoAtual.Z - pontoAnterior.Z) / comprimentoAnterior
                - (pontoPosterior.Z - pontoAtual.Z) / comprimentoPosterior;

            if (Math.Abs(deflexao) <= request.DeflexaoLimite)
                return ResultadoEtapa.AjusteConcluido("ajuste concluido na etapa de deflexao.");

            double sinalAjuste = deflexao > 0 ? -1 : 1;
            double novaElevacao = pontoAtual.Z + (sinalAjuste * request.PassoIncrementalAjusteCota);

            ResultadoAjustePonto resultadoAjuste = SetPiPointElevation(featureLine, indicePi, novaElevacao);
            if (!resultadoAjuste.Sucesso)
                return ResultadoEtapa.DaFalha(resultadoAjuste.TipoFalha, resultadoAjuste.Mensagem);

            houveAlteracao = true;
        }

        return ResultadoEtapa.NaoConvergiu(
            $"nao convergiu na etapa de deflexao apos {request.NumeroTentativasPorPonto} tentativas.");
    }

    /// <summary>
    /// Reproduz o ajuste iterativo pela superficie base do terceiro script Python.
    /// </summary>
    private static bool AjustarPorSuperficie(
        FeatureLine featureLine,
        TinSurface superficieBase,
        TerraplenagemFeatureLinesRequest request,
        ICollection<string> logs)
    {
        bool houveAlteracao = false;
        int ajustesFeitos = 0;

        while (ajustesFeitos < request.MaximoAjustesPorFeatureLine)
        {
            Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
            int quantidadePontosPi = pontosPi.Count;

            if (quantidadePontosPi < 2)
            {
                logs.Add($"'{featureLine.Name}': falha geometrica na etapa de superficie, a feature line possui menos de 2 PI points.");
                break;
            }

            int indiceMaximo = Math.Min(request.NumeroMaximoPontosPorFeatureLine, quantidadePontosPi - 1);
            int ultimoIndiceSegmento = indiceMaximo - 1;

            if (ultimoIndiceSegmento < 0)
                break;

            if (!TentarEncontrarSegmentoCritico(
                    pontosPi,
                    superficieBase,
                    request.ToleranciaAltaSuperficie,
                    ultimoIndiceSegmento,
                    out SegmentoCritico segmentoCritico))
            {
                logs.Add($"'{featureLine.Name}': nenhum segmento ficou fora da tolerancia da superficie.");
                break;
            }

            ResultadoEtapa resultadoSegmento = ProcessarSegmentoCritico(
                featureLine,
                superficieBase,
                request,
                segmentoCritico,
                ref houveAlteracao);

            ajustesFeitos++;
            logs.Add(
                $"'{featureLine.Name}': segmento {segmentoCritico.IndiceInicio}-{segmentoCritico.IndiceFim} - {resultadoSegmento.Mensagem}");

            // O script Dynamo faz apenas um ajuste critico por execucao
            // para reavaliar a feature line completa na proxima rodada.
            break;
        }

        return houveAlteracao;
    }

    /// <summary>
    /// Procura o segmento cujo ponto medio esta mais distante da superficie.
    /// </summary>
    private static bool TentarEncontrarSegmentoCritico(
        Point3dCollection pontosPi,
        TinSurface superficieBase,
        double toleranciaAlta,
        int ultimoIndiceSegmento,
        out SegmentoCritico segmentoCritico)
    {
        segmentoCritico = default;
        double moduloMaximo = 0;

        for (int indiceSegmento = 0; indiceSegmento <= ultimoIndiceSegmento; indiceSegmento++)
        {
            Point3d pontoInicial = pontosPi[indiceSegmento];
            Point3d pontoFinal = pontosPi[indiceSegmento + 1];

            if (!TentarObterDeltaMeioSegmento(pontoInicial, pontoFinal, superficieBase, out double deltaMeio))
                continue;

            double moduloAtual = Math.Abs(deltaMeio);
            if (moduloAtual <= toleranciaAlta || moduloAtual <= moduloMaximo)
                continue;

            moduloMaximo = moduloAtual;
            segmentoCritico = new SegmentoCritico(indiceSegmento, indiceSegmento + 1, deltaMeio);
        }

        return moduloMaximo > 0;
    }

    /// <summary>
    /// Processa o segmento critico escolhido na etapa de comparacao com a superficie.
    /// </summary>
    private static ResultadoEtapa ProcessarSegmentoCritico(
        FeatureLine featureLine,
        TinSurface superficieBase,
        TerraplenagemFeatureLinesRequest request,
        SegmentoCritico segmentoCritico,
        ref bool houveAlteracao)
    {
        Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);

        double? deltaPiInicial = ObterDeltaPonto(pontosPi[segmentoCritico.IndiceInicio], superficieBase);
        double? deltaPiFinal = ObterDeltaPonto(pontosPi[segmentoCritico.IndiceFim], superficieBase);
        IReadOnlyList<int> indicesAjuste = EscolherIndicesAjuste(
            segmentoCritico.IndiceInicio,
            segmentoCritico.IndiceFim,
            deltaPiInicial,
            deltaPiFinal,
            request.ToleranciaBaixaSuperficie);

        double? deltaMeioAnterior = null;
        int estagnado = 0;

        for (int tentativaLocal = 1; tentativaLocal <= request.NumeroTentativasPorPonto; tentativaLocal++)
        {
            pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);

            if (!TentarObterDeltaMeioSegmento(
                    pontosPi[segmentoCritico.IndiceInicio],
                    pontosPi[segmentoCritico.IndiceFim],
                    superficieBase,
                    out double deltaMeioAtual))
            {
                return ResultadoEtapa.FalhaGeometrica(
                    "falha geometrica na etapa de superficie: o ponto medio do segmento saiu da superficie base.");
            }

            if (Math.Abs(deltaMeioAtual) <= request.ToleranciaAltaSuperficie)
            {
                return ResultadoEtapa.AjusteConcluido(
                    $"ajuste concluido na etapa de superficie; delta inicial = {segmentoCritico.DeltaInicial.ToString("F4", CultureInfo.InvariantCulture)}.");
            }

            double deltaElevacao = deltaMeioAtual > 0
                ? -request.PassoIncrementalAjusteCota
                : request.PassoIncrementalAjusteCota;

            foreach (int indicePi in indicesAjuste)
            {
                pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
                double novaElevacao = pontosPi[indicePi].Z + deltaElevacao;

                ResultadoAjustePonto resultadoAjuste = SetPiPointElevation(featureLine, indicePi, novaElevacao);
                if (!resultadoAjuste.Sucesso)
                    return ResultadoEtapa.DaFalha(resultadoAjuste.TipoFalha, resultadoAjuste.Mensagem);

                houveAlteracao = true;
            }

            pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
            if (!TentarObterDeltaMeioSegmento(
                    pontosPi[segmentoCritico.IndiceInicio],
                    pontosPi[segmentoCritico.IndiceFim],
                    superficieBase,
                    out double deltaMeioNovo))
            {
                return ResultadoEtapa.FalhaGeometrica(
                    "falha geometrica na etapa de superficie: nao foi possivel recalcular o ponto medio do segmento apos o ajuste.");
            }

            if (deltaMeioAnterior.HasValue)
            {
                if (Math.Abs(deltaMeioNovo - deltaMeioAnterior.Value) < 1e-12)
                    estagnado++;
                else
                    estagnado = 0;
            }

            deltaMeioAnterior = deltaMeioNovo;

            if (estagnado >= 5)
            {
                return ResultadoEtapa.NaoConvergiu(
                    $"nao convergiu na etapa de superficie: estagnacao detectada apos {tentativaLocal} tentativas.");
            }
        }

        return ResultadoEtapa.NaoConvergiu(
            $"nao convergiu na etapa de superficie apos {request.NumeroTentativasPorPonto} tentativas.");
    }

    /// <summary>
    /// Calcula a distancia 2D entre dois pontos, ignorando a elevacao.
    /// </summary>
    private static double Distancia2d(Point3d pontoInicial, Point3d pontoFinal)
    {
        double deltaX = pontoFinal.X - pontoInicial.X;
        double deltaY = pontoFinal.Y - pontoInicial.Y;
        return Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
    }

    /// <summary>
    /// Calcula o delta entre a elevacao do PI e a elevacao da superficie no mesmo XY.
    /// </summary>
    private static double? ObterDeltaPonto(Point3d ponto, TinSurface superficieBase)
    {
        try
        {
            double elevacaoSuperficie = superficieBase.FindElevationAtXY(ponto.X, ponto.Y);
            return ponto.Z - elevacaoSuperficie;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Calcula o delta no meio do segmento, como o script Python faz.
    /// </summary>
    private static bool TentarObterDeltaMeioSegmento(
        Point3d pontoInicial,
        Point3d pontoFinal,
        TinSurface superficieBase,
        out double deltaMeio)
    {
        deltaMeio = 0;

        double xMeio = (pontoInicial.X + pontoFinal.X) * 0.5;
        double yMeio = (pontoInicial.Y + pontoFinal.Y) * 0.5;
        double zMeio = (pontoInicial.Z + pontoFinal.Z) * 0.5;

        try
        {
            double elevacaoSuperficie = superficieBase.FindElevationAtXY(xMeio, yMeio);
            deltaMeio = zMeio - elevacaoSuperficie;
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Decide se deve ajustar um ou dois PIs no segmento critico.
    /// </summary>
    private static IReadOnlyList<int> EscolherIndicesAjuste(
        int indicePiInicial,
        int indicePiFinal,
        double? deltaPiInicial,
        double? deltaPiFinal,
        double toleranciaBaixa)
    {
        double valorInicial = deltaPiInicial ?? 0;
        double valorFinal = deltaPiFinal ?? 0;

        bool inicialAlterado = Math.Abs(valorInicial) > toleranciaBaixa;
        bool finalAlterado = Math.Abs(valorFinal) > toleranciaBaixa;

        if (inicialAlterado && finalAlterado)
            return new[] { indicePiInicial, indicePiFinal };

        if (inicialAlterado)
            return new[] { indicePiInicial };

        if (finalAlterado)
            return new[] { indicePiFinal };

        return Math.Abs(valorInicial) <= Math.Abs(valorFinal)
            ? new[] { indicePiInicial }
            : new[] { indicePiFinal };
    }

    /// <summary>
    /// Ajusta a elevacao de um PI point convertendo o indice de PI para o indice de AllPoints.
    /// O ajuste falha explicitamente quando o mapeamento nao e confiavel.
    /// </summary>
    private static ResultadoAjustePonto SetPiPointElevation(
        FeatureLine featureLine,
        int indicePi,
        double novaElevacao)
    {
        Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
        Point3dCollection todosOsPontos = featureLine.GetPoints(FeatureLinePointType.AllPoints);

        if (indicePi < 0 || indicePi >= pontosPi.Count)
        {
            return ResultadoAjustePonto.Falha(
                TipoResultadoEtapa.FalhaMapeamento,
                "falha de mapeamento: o indice do PI point esta fora do intervalo esperado.");
        }

        Point3d pontoPi = pontosPi[indicePi];
        // O ajuste sempre acontece em AllPoints porque e esse indice que o Civil 3D aceita no SetPointElevation.
        if (!TryLocalizarIndiceAllPoints(
                todosOsPontos,
                pontoPi,
                out int indiceAllPoints,
                out string mensagemMapeamento))
        {
            return ResultadoAjustePonto.Falha(
                TipoResultadoEtapa.FalhaMapeamento,
                mensagemMapeamento);
        }

        try
        {
            featureLine.SetPointElevation(indiceAllPoints, novaElevacao);
            return ResultadoAjustePonto.ComSucesso();
        }
        catch (System.Exception ex)
        {
            return ResultadoAjustePonto.Falha(
                TipoResultadoEtapa.FalhaGeometrica,
                $"falha geometrica ao aplicar a nova elevacao: {ex.Message}");
        }
    }

    /// <summary>
    /// Localiza o indice correspondente do PI point em AllPoints apenas quando a correspondencia e confiavel.
    /// </summary>
    private static bool TryLocalizarIndiceAllPoints(
        Point3dCollection todosOsPontos,
        Point3d pontoPi,
        out int indiceAllPoints,
        out string mensagem)
    {
        indiceAllPoints = -1;
        mensagem = string.Empty;

        for (int indice = 0; indice < todosOsPontos.Count; indice++)
        {
            Point3d pontoAtual = todosOsPontos[indice];
            double distancia =
                Math.Sqrt(
                    Math.Pow(pontoAtual.X - pontoPi.X, 2)
                    + Math.Pow(pontoAtual.Y - pontoPi.Y, 2));

            if (distancia > ToleranciaMapeamentoXY)
                continue;

            indiceAllPoints = indice;
            return true;
        }

        mensagem =
            $"falha de mapeamento: nao foi possivel mapear o PI para AllPoints dentro da tolerancia XY de {ToleranciaMapeamentoXY.ToString("G", CultureInfo.InvariantCulture)}.";
        return false;
    }

    /// <summary>
    /// Estrutura imutavel com os dados preparados antes da escrita.
    /// </summary>
    private sealed record PreparacaoTerraplenagem(
        ObjectId IdSuperficieBase,
        int TotalFeatureLinesNoSite,
        List<ObjectId> IdsFeatureLinesFiltradas);

    /// <summary>
    /// Estrutura usada para representar um poligono de filtro encontrado na layer.
    /// </summary>
    private sealed record CandidatoPoligono(
        string Tipo,
        string Handle,
        List<Point2d> Vertices);

    /// <summary>
    /// Estrutura simples para devolver o resultado da selecao do poligono.
    /// </summary>
    private sealed record ResultadoSelecaoPoligono(
        bool Sucesso,
        string Mensagem,
        List<Point2d> Vertices)
    {
        public static ResultadoSelecaoPoligono Falha(string mensagem) =>
            new(false, mensagem, new List<Point2d>());

        public static ResultadoSelecaoPoligono SucessoCom(List<Point2d> vertices) =>
            new(true, string.Empty, vertices);
    }

    /// <summary>
    /// Estrutura que representa o segmento mais critico na etapa de superficie.
    /// </summary>
    private readonly record struct SegmentoCritico(
        int IndiceInicio,
        int IndiceFim,
        double DeltaInicial);

    /// <summary>
    /// Estrutura padronizada para retorno das etapas internas.
    /// </summary>
    private readonly record struct ResultadoEtapa(
        TipoResultadoEtapa Tipo,
        string Mensagem)
    {
        public static ResultadoEtapa AjusteConcluido(string mensagem) =>
            new(TipoResultadoEtapa.AjusteConcluido, mensagem);

        public static ResultadoEtapa FalhaGeometrica(string mensagem) =>
            new(TipoResultadoEtapa.FalhaGeometrica, mensagem);

        public static ResultadoEtapa NaoConvergiu(string mensagem) =>
            new(TipoResultadoEtapa.NaoConvergiu, mensagem);

        public static ResultadoEtapa DaFalha(TipoResultadoEtapa tipo, string mensagem) =>
            new(tipo, mensagem);
    }

    /// <summary>
    /// Estrutura de retorno do ajuste de um unico ponto.
    /// </summary>
    private readonly record struct ResultadoAjustePonto(
        bool Sucesso,
        TipoResultadoEtapa TipoFalha,
        string Mensagem)
    {
        public static ResultadoAjustePonto ComSucesso() =>
            new(true, TipoResultadoEtapa.AjusteConcluido, string.Empty);

        public static ResultadoAjustePonto Falha(TipoResultadoEtapa tipoFalha, string mensagem) =>
            new(false, tipoFalha, mensagem);
    }

    /// <summary>
    /// Categoria usada para separar sucesso, falhas e nao convergencia nos logs.
    /// </summary>
    private enum TipoResultadoEtapa
    {
        AjusteConcluido,
        FalhaGeometrica,
        FalhaMapeamento,
        NaoConvergiu
    }
}
