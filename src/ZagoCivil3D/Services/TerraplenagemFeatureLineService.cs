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
    /// Tolerancia usada para comparar coordenadas XY entre PI points e all points.
    /// </summary>
    private const double ToleranciaComparacaoXY = 1e-6;

    /// <summary>
    /// Executa o fluxo completo de terraplenagem dentro de uma unica transacao.
    /// </summary>
    public static TerraplenagemFeatureLinesResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        TerraplenagemFeatureLinesRequest request)
    {
        var resultado = new TerraplenagemFeatureLinesResultado();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        Site? site = ObterSitePorNome(civilDoc, transacao, request.NomeSite);
        if (site == null)
        {
            resultado.MensagensErro.Add($"Site '{request.NomeSite}' nao encontrado.");
            return resultado;
        }

        TinSurface? superficieBase = ObterSuperficiePorNome(civilDoc, transacao, request.NomeSuperficieBase);
        if (superficieBase == null)
        {
            resultado.MensagensErro.Add($"Superficie '{request.NomeSuperficieBase}' nao encontrada.");
            return resultado;
        }

        List<Point2d> verticesPoligono = ObterVerticesPoligonoDaCamada(db, transacao, request.NomeCamadaPoligono);
        if (verticesPoligono.Count < 3)
        {
            resultado.MensagensErro.Add(
                $"Nao foi encontrado um poligono fechado valido na layer '{request.NomeCamadaPoligono}'.");
            return resultado;
        }

        List<ObjectId> idsFeatureLines = site.GetFeatureLineIds().Cast<ObjectId>().ToList();
        resultado.TotalFeatureLinesNoSite = idsFeatureLines.Count;

        List<ObjectId> idsFiltrados = FiltrarFeatureLinesDentroPoligono(transacao, idsFeatureLines, verticesPoligono);
        resultado.TotalFeatureLinesFiltradas = idsFiltrados.Count;

        if (idsFiltrados.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhuma feature line do site ficou totalmente dentro do poligono informado.");
            return resultado;
        }

        foreach (ObjectId idFeatureLine in idsFiltrados)
        {
            if (transacao.GetObject(idFeatureLine, OpenMode.ForWrite) is not FeatureLine featureLine)
            {
                resultado.MensagensErro.Add($"Nao foi possivel abrir a feature line {idFeatureLine.Handle} para escrita.");
                continue;
            }

            try
            {
                bool ajustouDeflexao = AjustarDeflexao(featureLine, request, resultado.Logs);
                if (ajustouDeflexao)
                    resultado.TotalFeatureLinesComDeflexaoAjustada++;

                bool ajustouSuperficie = AjustarPorSuperficie(featureLine, superficieBase, request, resultado.Logs);
                if (ajustouSuperficie)
                    resultado.TotalFeatureLinesComSuperficieAjustada++;
            }
            catch (Exception ex)
            {
                resultado.MensagensErro.Add($"Erro ao processar a feature line '{featureLine.Name}': {ex.Message}");
            }
        }

        transacao.Commit();
        ed.WriteMessage("\n[ZagoCivil3D] Fluxo de terraplenagem concluido.");
        return resultado;
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
    /// Replica a etapa do Dynamo que pega o primeiro objeto da layer e usa sua geometria como filtro.
    /// </summary>
    private static List<Point2d> ObterVerticesPoligonoDaCamada(Database db, Transaction transacao, string nomeCamada)
    {
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

            if (vertices.Count >= 3)
                return vertices;
        }

        return new List<Point2d>();
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
    /// Mantem apenas as feature lines cujos pontos inicial e final estao dentro do poligono.
    /// </summary>
    private static List<ObjectId> FiltrarFeatureLinesDentroPoligono(
        Transaction transacao,
        IReadOnlyCollection<ObjectId> idsFeatureLines,
        IReadOnlyList<Point2d> verticesPoligono)
    {
        var filtradas = new List<ObjectId>();

        foreach (ObjectId idFeatureLine in idsFeatureLines)
        {
            if (transacao.GetObject(idFeatureLine, OpenMode.ForRead) is not FeatureLine featureLine)
                continue;

            Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
            if (pontosPi.Count < 2)
                continue;

            Point3d pontoInicial = pontosPi[0];
            Point3d pontoFinal = pontosPi[pontosPi.Count - 1];

            bool dentroInicial = PontoDentroPoligono(new Point2d(pontoInicial.X, pontoInicial.Y), verticesPoligono);
            bool dentroFinal = PontoDentroPoligono(new Point2d(pontoFinal.X, pontoFinal.Y), verticesPoligono);

            if (dentroInicial && dentroFinal)
                filtradas.Add(idFeatureLine);
        }

        return filtradas;
    }

    /// <summary>
    /// Implementa o algoritmo de ray casting usado no script Python do Dynamo.
    /// </summary>
    private static bool PontoDentroPoligono(Point2d ponto, IReadOnlyList<Point2d> poligono)
    {
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
                logs.Add($"'{featureLine.Name}': menos de 3 PI points, etapa de deflexao ignorada.");
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
                bool pontoDentroDoLimite = false;

                for (int tentativaLocal = 1; tentativaLocal <= request.NumeroTentativasPorPonto; tentativaLocal++)
                {
                    pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);

                    Point3d pontoAnterior = pontosPi[indicePi - 1];
                    Point3d pontoAtual = pontosPi[indicePi];
                    Point3d pontoPosterior = pontosPi[indicePi + 1];

                    double comprimentoAnterior = Distancia2d(pontoAnterior, pontoAtual);
                    double comprimentoPosterior = Distancia2d(pontoAtual, pontoPosterior);

                    if (comprimentoAnterior <= ToleranciaComparacaoXY || comprimentoPosterior <= ToleranciaComparacaoXY)
                    {
                        logs.Add($"'{featureLine.Name}': PI {indicePi} ignorado por comprimento 2D nulo.");
                        break;
                    }

                    double deflexao =
                        (pontoAtual.Z - pontoAnterior.Z) / comprimentoAnterior
                        - (pontoPosterior.Z - pontoAtual.Z) / comprimentoPosterior;

                    if (Math.Abs(deflexao) <= request.DeflexaoLimite)
                    {
                        pontosOk++;
                        pontoDentroDoLimite = true;
                        break;
                    }

                    double sinalAjuste = deflexao > 0 ? -1 : 1;
                    double novaElevacao = pontoAtual.Z + (sinalAjuste * request.PassoIncrementalAjusteCota);

                    if (!SetPiPointElevation(featureLine, indicePi, novaElevacao))
                    {
                        logs.Add($"'{featureLine.Name}': falha ao ajustar o PI {indicePi} na etapa de deflexao.");
                        break;
                    }

                    houveAlteracao = true;
                }

                if (!pontoDentroDoLimite)
                    logs.Add($"'{featureLine.Name}': PI {indicePi} nao convergiu na etapa de deflexao.");
            }

            double percentualOk = totalPontos == 0 ? 1 : (double)pontosOk / totalPontos;
            logs.Add(
                $"'{featureLine.Name}': passada {passadaGlobal}, pontos OK = {pontosOk}/{totalPontos}, percentual = {percentualOk.ToString("P2", CultureInfo.InvariantCulture)}.");

            if (percentualOk >= request.PercentualObjetivo)
                break;
        }

        return houveAlteracao;
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
                logs.Add($"'{featureLine.Name}': menos de 2 PI points, etapa de superficie ignorada.");
                break;
            }

            int indiceMaximo = Math.Min(request.NumeroMaximoPontosPorFeatureLine, quantidadePontosPi - 1);
            int ultimoIndiceSegmento = indiceMaximo - 1;

            if (ultimoIndiceSegmento < 0)
                break;

            (int Inicio, int Fim)? segmentoCritico = null;
            double deltaCritico = 0;
            double moduloMaximo = 0;

            for (int indiceSegmento = 0; indiceSegmento <= ultimoIndiceSegmento; indiceSegmento++)
            {
                Point3d pontoInicial = pontosPi[indiceSegmento];
                Point3d pontoFinal = pontosPi[indiceSegmento + 1];

                if (!TentarObterDeltaMeioSegmento(pontoInicial, pontoFinal, superficieBase, out double deltaMeio))
                    continue;

                double moduloAtual = Math.Abs(deltaMeio);
                if (moduloAtual > request.ToleranciaAltaSuperficie && moduloAtual > moduloMaximo)
                {
                    moduloMaximo = moduloAtual;
                    deltaCritico = deltaMeio;
                    segmentoCritico = (indiceSegmento, indiceSegmento + 1);
                }
            }

            if (segmentoCritico == null)
            {
                logs.Add($"'{featureLine.Name}': nenhum segmento fora da tolerancia de superficie.");
                break;
            }

            int indiceInicio = segmentoCritico.Value.Inicio;
            int indiceFim = segmentoCritico.Value.Fim;

            double? deltaPiInicial = ObterDeltaPonto(pontosPi[indiceInicio], superficieBase);
            double? deltaPiFinal = ObterDeltaPonto(pontosPi[indiceFim], superficieBase);
            IReadOnlyList<int> indicesAjuste = EscolherIndicesAjuste(
                indiceInicio,
                indiceFim,
                deltaPiInicial,
                deltaPiFinal,
                request.ToleranciaBaixaSuperficie);

            double? deltaMeioAnterior = null;
            int estagnado = 0;

            for (int tentativaLocal = 1; tentativaLocal <= request.NumeroTentativasPorPonto; tentativaLocal++)
            {
                pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);

                if (!TentarObterDeltaMeioSegmento(pontosPi[indiceInicio], pontosPi[indiceFim], superficieBase, out double deltaMeioAtual))
                {
                    logs.Add($"'{featureLine.Name}': meio do segmento critico saiu da superficie.");
                    break;
                }

                if (Math.Abs(deltaMeioAtual) <= request.ToleranciaAltaSuperficie)
                {
                    logs.Add(
                        $"'{featureLine.Name}': segmento {indiceInicio}-{indiceFim} entrou na tolerancia de superficie.");
                    break;
                }

                double deltaElevacao = deltaMeioAtual > 0
                    ? -request.PassoIncrementalAjusteCota
                    : request.PassoIncrementalAjusteCota;

                foreach (int indicePi in indicesAjuste)
                {
                    pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
                    double novaElevacao = pontosPi[indicePi].Z + deltaElevacao;

                    if (!SetPiPointElevation(featureLine, indicePi, novaElevacao))
                    {
                        logs.Add($"'{featureLine.Name}': falha ao ajustar o PI {indicePi} na etapa de superficie.");
                        break;
                    }

                    houveAlteracao = true;
                }

                pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
                if (!TentarObterDeltaMeioSegmento(pontosPi[indiceInicio], pontosPi[indiceFim], superficieBase, out double deltaMeioNovo))
                    break;

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
                    logs.Add($"'{featureLine.Name}': estagnacao detectada no segmento {indiceInicio}-{indiceFim}.");
                    break;
                }
            }

            ajustesFeitos++;
            logs.Add(
                $"'{featureLine.Name}': ajuste por superficie concluido no segmento {indiceInicio}-{indiceFim}, delta inicial = {deltaCritico.ToString("F4", CultureInfo.InvariantCulture)}.");

            // O script Dynamo faz apenas um ajuste critico por execucao
            // para reavaliar a feature line completa na proxima rodada.
            break;
        }

        return houveAlteracao;
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
    /// Ajusta a elevacao de um PI point convertendo o indice de PI para o indice de all points.
    /// </summary>
    private static bool SetPiPointElevation(FeatureLine featureLine, int indicePi, double novaElevacao)
    {
        Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
        Point3dCollection todosOsPontos = featureLine.GetPoints(FeatureLinePointType.AllPoints);

        if (indicePi < 0 || indicePi >= pontosPi.Count)
            return false;

        Point3d pontoPi = pontosPi[indicePi];
        int indiceAllPoints = LocalizarIndiceAllPoints(todosOsPontos, pontoPi);

        if (indiceAllPoints < 0)
            return false;

        featureLine.SetPointElevation(indiceAllPoints, novaElevacao);
        return true;
    }

    /// <summary>
    /// Procura o indice do ponto equivalente dentro da colecao AllPoints.
    /// </summary>
    private static int LocalizarIndiceAllPoints(Point3dCollection todosOsPontos, Point3d pontoPi)
    {
        int indiceMaisProximo = -1;
        double menorDistancia = double.MaxValue;

        for (int indice = 0; indice < todosOsPontos.Count; indice++)
        {
            Point3d pontoAtual = todosOsPontos[indice];
            double distancia =
                Math.Sqrt(
                    Math.Pow(pontoAtual.X - pontoPi.X, 2)
                    + Math.Pow(pontoAtual.Y - pontoPi.Y, 2));

            if (distancia < menorDistancia)
            {
                menorDistancia = distancia;
                indiceMaisProximo = indice;
            }

            if (distancia <= ToleranciaComparacaoXY)
                return indice;
        }

        return indiceMaisProximo;
    }
}
