using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

/// <summary>
/// Logica que replica o Dynamo "CRIAR TALVEGUE DOS CATCHMENTS A PARTIR DAS
/// POLILINHAS": para cada polyline desenhada em uma layer especifica, encontra
/// o catchment cujo boundary contem os pontos inicial e final da polyline e
/// chama Catchment.SetFlowPath() usando os vertices da polyline.
/// </summary>
public static class TalveguesCatchmentsService
{
    /// <summary>
    /// Executa o fluxo completo. A transacao envolve leitura e escrita porque
    /// Catchment.SetFlowPath precisa ocorrer enquanto o catchment esta aberto
    /// para escrita.
    /// </summary>
    public static CriarTalveguesCatchmentsResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        CriarTalveguesCatchmentsRequest request)
    {
        var resultado = new CriarTalveguesCatchmentsResultado();

        if (string.IsNullOrWhiteSpace(request.NomeCamadaPolilinhas))
        {
            resultado.MensagensErro.Add("Layer das polilinhas nao informada.");
            return resultado;
        }

        using Transaction transacao = db.TransactionManager.StartTransaction();

        List<PolylineAnalisada> polilinhas = ColetarPolilinhasDaCamada(
            db, transacao, request.NomeCamadaPolilinhas, resultado);

        resultado.TotalPolilinhas = polilinhas.Count;

        if (polilinhas.Count == 0)
        {
            resultado.MensagensErro.Add(
                $"Nenhuma polyline valida encontrada na layer '{request.NomeCamadaPolilinhas}'.");
            transacao.Commit();
            return resultado;
        }

        List<CatchmentAnalisado> catchments = ColetarCatchments(db, transacao, resultado);
        resultado.TotalCatchments = catchments.Count;

        if (catchments.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhum catchment encontrado no desenho.");
            transacao.Commit();
            return resultado;
        }

        ed.WriteMessage(
            $"\n[ZagoCivil3D] Polilinhas: {polilinhas.Count} | Catchments: {catchments.Count}.");

        // Conjunto de catchments ja atualizados neste run — evita que duas
        // polilinhas disputem o mesmo catchment sem sinalizar ao usuario.
        var catchmentsAtualizados = new HashSet<ObjectId>();

        foreach (PolylineAnalisada polilinha in polilinhas)
        {
            ProcessarPolilinha(
                transacao,
                polilinha,
                catchments,
                catchmentsAtualizados,
                request,
                resultado);
        }

        transacao.Commit();

        ed.WriteMessage("\n[ZagoCivil3D] Fluxo de talvegues concluido.");
        return resultado;
    }

    /// <summary>
    /// Coleta as polilinhas do ModelSpace que estao na layer informada e
    /// possuem ao menos dois vertices.
    /// </summary>
    private static List<PolylineAnalisada> ColetarPolilinhasDaCamada(
        Database db,
        Transaction transacao,
        string nomeCamada,
        CriarTalveguesCatchmentsResultado resultado)
    {
        var saida = new List<PolylineAnalisada>();
        var tabelaBlocos = (BlockTable)transacao.GetObject(db.BlockTableId, OpenMode.ForRead);
        var espacoModelo =
            (BlockTableRecord)transacao.GetObject(tabelaBlocos[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        foreach (ObjectId idEntidade in espacoModelo)
        {
            if (transacao.GetObject(idEntidade, OpenMode.ForRead) is not Autodesk.AutoCAD.DatabaseServices.Entity entidade)
                continue;

            if (!string.Equals(entidade.Layer, nomeCamada, StringComparison.OrdinalIgnoreCase))
                continue;

            List<Point3d> vertices3d = entidade switch
            {
                Polyline polyline => ExtrairVerticesPolyline(polyline),
                Polyline2d polyline2d => ExtrairVerticesPolyline2d(polyline2d, transacao),
                Polyline3d polyline3d => ExtrairVerticesPolyline3d(polyline3d, transacao),
                _ => new List<Point3d>()
            };

            if (vertices3d.Count < 2)
            {
                resultado.Logs.Add(
                    $"Entidade {entidade.Handle} na layer '{nomeCamada}' ignorada (menos de 2 vertices).");
                continue;
            }

            saida.Add(new PolylineAnalisada(entidade.Handle.ToString(), vertices3d));
        }

        return saida;
    }

    /// <summary>
    /// Le todos os catchments do documento Civil 3D via CatchmentGroupCollection.
    /// Preserva o nome e o boundary 2D para uso no containment test.
    /// </summary>
    private static List<CatchmentAnalisado> ColetarCatchments(
        Database db,
        Transaction transacao,
        CriarTalveguesCatchmentsResultado resultado)
    {
        var saida = new List<CatchmentAnalisado>();
        CatchmentGroupCollection grupos = CatchmentGroupCollection.GetCatchmentGroups(db);

        foreach (ObjectId idGrupo in grupos)
        {
            if (idGrupo.IsNull || idGrupo.IsErased) continue;

            if (transacao.GetObject(idGrupo, OpenMode.ForRead) is not CatchmentGroup grupo)
                continue;

            ObjectIdCollection idsCatchments = grupo.GetAllCatchmentIds();
            foreach (ObjectId idCatchment in idsCatchments)
            {
                if (idCatchment.IsNull || idCatchment.IsErased) continue;

                try
                {
                    if (transacao.GetObject(idCatchment, OpenMode.ForRead) is not Catchment catchment)
                        continue;

                    Point2dCollection boundary2d = catchment.BoundaryPolyline2d;
                    if (boundary2d == null || boundary2d.Count < 3)
                    {
                        resultado.Logs.Add($"Catchment '{catchment.Name}' ignorado (boundary com menos de 3 vertices).");
                        continue;
                    }

                    var vertices = new List<Point2d>();
                    foreach (Point2d p in boundary2d)
                        vertices.Add(p);

                    // O FlowPath existente e lido aqui para decidir depois se substituimos.
                    bool temFlowPath = false;
                    try
                    {
                        FlowPath caminho = catchment.FlowPath;
                        temFlowPath = caminho != null && caminho.FlowSegmentCount > 0;
                    }
                    catch
                    {
                        temFlowPath = false;
                    }

                    saida.Add(new CatchmentAnalisado(
                        idCatchment,
                        catchment.Name,
                        vertices,
                        temFlowPath));
                }
                catch (Exception ex)
                {
                    resultado.Logs.Add($"Catchment {idCatchment.Handle} ignorado: {ex.Message}");
                }
            }
        }

        return saida;
    }

    /// <summary>
    /// Associa uma polyline a um catchment verificando se o primeiro e o ultimo
    /// vertice da polyline estao dentro do boundary do catchment. Quando o
    /// match acontece, chama Catchment.SetFlowPath com os vertices da polyline.
    /// </summary>
    private static void ProcessarPolilinha(
        Transaction transacao,
        PolylineAnalisada polilinha,
        IReadOnlyList<CatchmentAnalisado> catchments,
        HashSet<ObjectId> catchmentsAtualizados,
        CriarTalveguesCatchmentsRequest request,
        CriarTalveguesCatchmentsResultado resultado)
    {
        Point3d pontoInicial = polilinha.Vertices[0];
        Point3d pontoFinal = polilinha.Vertices[^1];
        var ponto2dInicial = new Point2d(pontoInicial.X, pontoInicial.Y);
        var ponto2dFinal = new Point2d(pontoFinal.X, pontoFinal.Y);

        CatchmentAnalisado? correspondente = catchments.FirstOrDefault(c =>
            PontoDentroPoligono(ponto2dInicial, c.BoundaryXY, request.ToleranciaBorda)
            && PontoDentroPoligono(ponto2dFinal, c.BoundaryXY, request.ToleranciaBorda));

        if (correspondente == null)
        {
            resultado.TotalPolilinhasSemCatchment++;
            resultado.Logs.Add(
                $"Polyline {polilinha.Handle} nao esta contida em nenhum catchment (revisar).");
            return;
        }

        if (catchmentsAtualizados.Contains(correspondente.Id))
        {
            resultado.Logs.Add(
                $"Polyline {polilinha.Handle} ignorada: catchment '{correspondente.Name}' ja foi atualizado por outra polyline nesta execucao.");
            return;
        }

        if (correspondente.TemFlowPathExistente && !request.SubstituirFlowPathExistente)
        {
            resultado.TotalCatchmentsIgnorados++;
            resultado.Logs.Add(
                $"Catchment '{correspondente.Name}' ja possui FlowPath e a opcao de substituir esta desligada.");
            return;
        }

        try
        {
            if (transacao.GetObject(correspondente.Id, OpenMode.ForWrite) is not Catchment catchment)
            {
                resultado.MensagensErro.Add(
                    $"Nao foi possivel abrir o catchment '{correspondente.Name}' para escrita.");
                return;
            }

            var pontos3d = new Point3dCollection();
            foreach (Point3d ponto in polilinha.Vertices)
                pontos3d.Add(ponto);

            // A API nova Civil 3D: Catchment.SetFlowPath(Point3dCollection)
            // esta marcada como obsoleta e a substituta e FlowPath.SetPath.
            FlowPath caminhoFluxo = catchment.GetFlowPath();
            caminhoFluxo.SetPath(pontos3d);

            catchmentsAtualizados.Add(correspondente.Id);
            resultado.NomesCatchmentsAtualizados.Add(correspondente.Name);

            if (correspondente.TemFlowPathExistente)
                resultado.TotalFlowPathsSubstituidos++;
            else
                resultado.TotalFlowPathsCriados++;

            resultado.Logs.Add(
                $"Polyline {polilinha.Handle} -> Catchment '{correspondente.Name}' ({polilinha.Vertices.Count} vertices).");
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add(
                $"Erro ao atualizar catchment '{correspondente.Name}': {ex.Message}");
        }
    }

    /// <summary>
    /// Extrai vertices de uma polyline 2D padrao (LWPOLYLINE). As elevacoes
    /// sao mantidas apenas se a polyline tiver elevacao nao nula.
    /// </summary>
    private static List<Point3d> ExtrairVerticesPolyline(Polyline polyline)
    {
        var vertices = new List<Point3d>();
        for (int indice = 0; indice < polyline.NumberOfVertices; indice++)
        {
            Point2d ponto2d = polyline.GetPoint2dAt(indice);
            vertices.Add(new Point3d(ponto2d.X, ponto2d.Y, polyline.Elevation));
        }
        return vertices;
    }

    /// <summary>Extrai vertices de uma POLYLINE2D heavy.</summary>
    private static List<Point3d> ExtrairVerticesPolyline2d(Polyline2d polyline, Transaction transacao)
    {
        var vertices = new List<Point3d>();
        foreach (ObjectId idVertice in polyline)
        {
            if (transacao.GetObject(idVertice, OpenMode.ForRead) is Vertex2d vertice)
                vertices.Add(vertice.Position);
        }
        return vertices;
    }

    /// <summary>Extrai vertices de uma POLYLINE3D.</summary>
    private static List<Point3d> ExtrairVerticesPolyline3d(Polyline3d polyline, Transaction transacao)
    {
        var vertices = new List<Point3d>();
        foreach (ObjectId idVertice in polyline)
        {
            if (transacao.GetObject(idVertice, OpenMode.ForRead) is PolylineVertex3d vertice)
                vertices.Add(vertice.Position);
        }
        return vertices;
    }

    /// <summary>
    /// Algoritmo ray casting com tratamento de pontos sobre a borda. Espelha
    /// o mesmo criterio usado em TerraplenagemFeatureLineService para manter
    /// consistencia entre os comandos do plugin.
    /// </summary>
    private static bool PontoDentroPoligono(
        Point2d ponto,
        IReadOnlyList<Point2d> poligono,
        double toleranciaBorda)
    {
        for (int indice = 0; indice < poligono.Count; indice++)
        {
            Point2d inicio = poligono[indice];
            Point2d fim = poligono[(indice + 1) % poligono.Count];

            if (PontoSobreSegmento(ponto, inicio, fim, toleranciaBorda))
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

    /// <summary>Verifica se um ponto esta sobre uma aresta do poligono.</summary>
    private static bool PontoSobreSegmento(
        Point2d ponto,
        Point2d inicio,
        Point2d fim,
        double tolerancia)
    {
        double areaDobrada =
            ((fim.X - inicio.X) * (ponto.Y - inicio.Y))
            - ((fim.Y - inicio.Y) * (ponto.X - inicio.X));

        if (Math.Abs(areaDobrada) > tolerancia)
            return false;

        double minX = Math.Min(inicio.X, fim.X) - tolerancia;
        double maxX = Math.Max(inicio.X, fim.X) + tolerancia;
        double minY = Math.Min(inicio.Y, fim.Y) - tolerancia;
        double maxY = Math.Max(inicio.Y, fim.Y) + tolerancia;

        return ponto.X >= minX
            && ponto.X <= maxX
            && ponto.Y >= minY
            && ponto.Y <= maxY;
    }

    /// <summary>Dados de uma polyline selecionada para processamento.</summary>
    private sealed record PolylineAnalisada(string Handle, List<Point3d> Vertices);

    /// <summary>Dados de um catchment indexados em memoria para o matching.</summary>
    private sealed record CatchmentAnalisado(
        ObjectId Id,
        string Name,
        List<Point2d> BoundaryXY,
        bool TemFlowPathExistente);
}
