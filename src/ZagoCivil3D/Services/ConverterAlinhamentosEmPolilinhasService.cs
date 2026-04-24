using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Logica de conversao de alinhamentos (Civil 3D <see cref="Alignment"/>) em
/// polilinhas 2D do AutoCAD (<see cref="Polyline"/>).
///
/// Linhas viram segmentos retos; arcos circulares sao preservados como
/// vertices com <c>bulge</c>; espirais (clotoides) sao discretizadas em
/// segmentos retos pelo passo configurado, usando
/// <see cref="Alignment.PointLocation(double, double, ref double, ref double)"/>.
/// </summary>
public static class ConverterAlinhamentosEmPolilinhasService
{
    /// <summary>
    /// Executa o fluxo completo de conversao.
    /// </summary>
    public static ConverterAlinhamentosEmPolilinhasResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        ConverterAlinhamentosEmPolilinhasRequest request)
    {
        var resultado = new ConverterAlinhamentosEmPolilinhasResultado();

        ObjectIdCollection idsAlinhamentos = civilDoc.GetAlignmentIds();
        resultado.TotalAlinhamentos = idsAlinhamentos.Count;

        if (idsAlinhamentos.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhum alinhamento encontrado no desenho.");
            return resultado;
        }

        string filtro = (request.FiltroNome ?? string.Empty).Trim();
        bool temFiltro = !string.IsNullOrEmpty(filtro);

        if (temFiltro)
        {
            ed.WriteMessage(
                $"\n[ZagoCivil3D] Filtrando alinhamentos que contem '{filtro}' no nome...");
        }
        else
        {
            ed.WriteMessage(
                $"\n[ZagoCivil3D] Convertendo {idsAlinhamentos.Count} alinhamento(s) em polilinhas 2D...");
        }

        // Em dry-run nao alteramos o desenho: a layer so e criada na execucao
        // real, evitando que a mera visualizacao do preview persista uma layer
        // nova no DWG.
        if (!request.DryRun)
            GarantirLayerExiste(db, request.NomeLayer, resultado);

        double passo = request.PassoDiscretizacaoEspirais <= 0
            ? 1.0
            : request.PassoDiscretizacaoEspirais;

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            var blockTable = (BlockTable)transacao.GetObject(db.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transacao.GetObject(
                blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            var idsParaApagar = new List<ObjectId>();

            foreach (ObjectId idAlinhamento in idsAlinhamentos)
            {
                try
                {
                    if (transacao.GetObject(idAlinhamento, OpenMode.ForRead) is not Alignment alinhamento)
                        continue;

                    string nomeAlinhamento = alinhamento.Name ?? "(sem nome)";

                    if (temFiltro && nomeAlinhamento.IndexOf(filtro, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        resultado.Logs.Add($"Alinhamento '{nomeAlinhamento}' ignorado pelo filtro '{filtro}'.");
                        continue;
                    }

                    resultado.TotalAlinhamentosFiltrados++;

                    var detalhe = new AlinhamentoConvertido
                    {
                        NomeAlinhamento = nomeAlinhamento,
                        ComprimentoAlinhamento = alinhamento.EndingStation - alinhamento.StartingStation
                    };

                    Polyline polilinha = ConstruirPolilinha(
                        alinhamento, request, passo, detalhe, resultado);

                    if (polilinha.NumberOfVertices < 2)
                    {
                        polilinha.Dispose();
                        resultado.Avisos.Add(
                            $"Alinhamento '{detalhe.NomeAlinhamento}' gerou polilinha com menos de 2 vertices. Ignorado.");
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(request.NomeLayer))
                        polilinha.Layer = request.NomeLayer;

                    polilinha.Elevation = request.Elevation;

                    // Captura antes do Dispose para que o dry-run reporte a mesma
                    // contagem de vertices que a execucao real criaria — segmentos
                    // conectados compartilham extremidades, entao N sub-entidades
                    // nao implicam N vertices.
                    detalhe.TotalVertices = polilinha.NumberOfVertices;

                    if (request.DryRun)
                    {
                        polilinha.Dispose();
                        resultado.Convertidos.Add(detalhe);
                        continue;
                    }

                    modelSpace.AppendEntity(polilinha);
                    transacao.AddNewlyCreatedDBObject(polilinha, true);

                    resultado.Convertidos.Add(detalhe);
                    resultado.TotalPolilinhasCriadas++;

                    if (request.ApagarAlinhamentosOriginais)
                        idsParaApagar.Add(idAlinhamento);
                }
                catch (Exception ex)
                {
                    resultado.MensagensErro.Add(
                        $"Erro ao converter alinhamento {idAlinhamento.Handle}: {ex.Message}");
                }
            }

            if (!request.DryRun && request.ApagarAlinhamentosOriginais)
            {
                foreach (ObjectId id in idsParaApagar)
                {
                    try
                    {
                        if (transacao.GetObject(id, OpenMode.ForWrite) is Alignment alinhamento)
                        {
                            alinhamento.Erase();
                            resultado.TotalAlinhamentosApagados++;
                        }
                    }
                    catch (Exception ex)
                    {
                        resultado.Logs.Add($"Nao foi possivel apagar alinhamento {id.Handle}: {ex.Message}");
                    }
                }
            }

            transacao.Commit();
        }
        catch (Exception ex)
        {
            try { transacao.Abort(); } catch { }
            resultado.MensagensErro.Add($"Erro geral na conversao: {ex.Message}");
        }

        if (temFiltro && resultado.TotalAlinhamentosFiltrados == 0)
        {
            resultado.MensagensErro.Add(
                $"Nenhum alinhamento contem '{filtro}' no nome. Verifique o filtro.");
        }

        return resultado;
    }

    /// <summary>
    /// Percorre as sub-entidades do alinhamento e constroi a polyline 2D
    /// correspondente, adicionando vertices com bulge para arcos e vertices
    /// discretizados por passo para espirais.
    /// </summary>
    private static Polyline ConstruirPolilinha(
        Alignment alinhamento,
        ConverterAlinhamentosEmPolilinhasRequest request,
        double passo,
        AlinhamentoConvertido detalhe,
        ConverterAlinhamentosEmPolilinhasResultado resultado)
    {
        var polilinha = new Polyline();
        polilinha.SetDatabaseDefaults();

        int indice = 0;
        Point2d? ultimoPonto = null;

        foreach (AlignmentEntity entidade in alinhamento.Entities)
        {
            for (int i = 0; i < entidade.SubEntityCount; i++)
            {
                AlignmentSubEntity sub = entidade[i];

                switch (sub.SubEntityType)
                {
                    case AlignmentSubEntityType.Line:
                        var linha = (AlignmentSubEntityLine)sub;
                        AdicionarSegmentoReto(
                            polilinha, ref indice, ref ultimoPonto,
                            linha.StartPoint, linha.EndPoint, bulge: 0);
                        detalhe.TotalLinhas++;
                        break;

                    case AlignmentSubEntityType.Arc:
                        var arco = (AlignmentSubEntityArc)sub;
                        if (request.PreservarArcos)
                        {
                            double bulge = CalcularBulge(arco);
                            AdicionarSegmentoArco(
                                polilinha, ref indice, ref ultimoPonto,
                                arco.StartPoint, arco.EndPoint, bulge);
                        }
                        else
                        {
                            DiscretizarPorEstacao(
                                alinhamento, polilinha, ref indice, ref ultimoPonto,
                                arco.StartStation, arco.EndStation, passo);
                        }
                        detalhe.TotalArcos++;
                        break;

                    case AlignmentSubEntityType.Spiral:
                        var espiral = (AlignmentSubEntitySpiral)sub;
                        int antes = indice;
                        DiscretizarPorEstacao(
                            alinhamento, polilinha, ref indice, ref ultimoPonto,
                            espiral.StartStation, espiral.EndStation, passo);
                        detalhe.VerticesDiscretizadosEspirais += indice - antes;
                        detalhe.TotalEspirais++;
                        break;

                    default:
                        resultado.Avisos.Add(
                            $"Sub-entidade desconhecida em '{alinhamento.Name}': {sub.SubEntityType}. Ignorada.");
                        break;
                }
            }
        }

        return polilinha;
    }

    /// <summary>
    /// Adiciona um segmento reto na polyline, garantindo que o ponto inicial
    /// seja inserido apenas quando diferente do ultimo ponto ja presente.
    /// </summary>
    private static void AdicionarSegmentoReto(
        Polyline polilinha,
        ref int indice,
        ref Point2d? ultimoPonto,
        Point2d inicio,
        Point2d fim,
        double bulge)
    {
        if (ultimoPonto == null || !PontosIguais(ultimoPonto.Value, inicio))
        {
            polilinha.AddVertexAt(indice, inicio, bulge, 0, 0);
            indice++;
        }
        else
        {
            polilinha.SetBulgeAt(indice - 1, bulge);
        }

        polilinha.AddVertexAt(indice, fim, 0, 0, 0);
        indice++;
        ultimoPonto = fim;
    }

    /// <summary>
    /// Adiciona um segmento com arco (bulge no vertice inicial).
    /// </summary>
    private static void AdicionarSegmentoArco(
        Polyline polilinha,
        ref int indice,
        ref Point2d? ultimoPonto,
        Point2d inicio,
        Point2d fim,
        double bulge)
    {
        AdicionarSegmentoReto(polilinha, ref indice, ref ultimoPonto, inicio, fim, bulge);
    }

    /// <summary>
    /// Discretiza um trecho do alinhamento pela estacao, adicionando vertices
    /// retos de <paramref name="estacaoInicio"/> ate <paramref name="estacaoFim"/>
    /// com passo <paramref name="passo"/>. O ultimo vertice corresponde
    /// exatamente a estacao final para preservar o fechamento do trecho.
    /// </summary>
    private static void DiscretizarPorEstacao(
        Alignment alinhamento,
        Polyline polilinha,
        ref int indice,
        ref Point2d? ultimoPonto,
        double estacaoInicio,
        double estacaoFim,
        double passo)
    {
        double comprimento = estacaoFim - estacaoInicio;
        if (comprimento <= 1e-9)
            return;

        int subdivisoes = Math.Max(1, (int)Math.Ceiling(comprimento / passo));
        double passoReal = comprimento / subdivisoes;

        for (int s = 0; s <= subdivisoes; s++)
        {
            double estacao = estacaoInicio + s * passoReal;
            if (s == subdivisoes)
                estacao = estacaoFim;

            Point2d ponto = ObterPontoEmEstacao(alinhamento, estacao);

            if (ultimoPonto != null && PontosIguais(ultimoPonto.Value, ponto))
                continue;

            polilinha.AddVertexAt(indice, ponto, 0, 0, 0);
            indice++;
            ultimoPonto = ponto;
        }
    }

    /// <summary>
    /// Calcula o <c>bulge</c> de um arco circular a partir do angulo central
    /// (<c>Delta</c>) e do sentido de curvatura.
    /// <c>bulge = tan(Delta/4)</c>, positivo para sentido anti-horario e
    /// negativo para sentido horario.
    /// </summary>
    private static double CalcularBulge(AlignmentSubEntityArc arco)
    {
        double delta = arco.Delta;

        if (delta <= 0 && arco.Radius > 0 && arco.Length > 0)
            delta = arco.Length / arco.Radius;

        double bulge = Math.Tan(delta / 4.0);

        if (arco.Clockwise)
            bulge = -bulge;

        return bulge;
    }

    /// <summary>
    /// Obtem o ponto (X,Y) no alinhamento para uma estacao dada, usando
    /// <see cref="Alignment.PointLocation"/>. Parametros sao passados por
    /// referencia conforme a assinatura da API do Civil 3D.
    /// </summary>
    private static Point2d ObterPontoEmEstacao(Alignment alinhamento, double estacao)
    {
        double easting = 0;
        double northing = 0;
        alinhamento.PointLocation(estacao, 0.0, ref easting, ref northing);
        return new Point2d(easting, northing);
    }

    private static bool PontosIguais(Point2d a, Point2d b)
    {
        const double tolerancia = 1e-6;
        double dx = a.X - b.X;
        double dy = a.Y - b.Y;
        return (dx * dx + dy * dy) <= tolerancia * tolerancia;
    }

    /// <summary>
    /// Garante que a layer destino exista. Se nao existir, cria-a em uma
    /// transacao propria antes da transacao principal de criacao das
    /// polilinhas.
    /// </summary>
    private static void GarantirLayerExiste(
        Database db,
        string nomeLayer,
        ConverterAlinhamentosEmPolilinhasResultado resultado)
    {
        if (string.IsNullOrWhiteSpace(nomeLayer))
            return;

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            var tabela = (LayerTable)transacao.GetObject(db.LayerTableId, OpenMode.ForRead);

            if (tabela.Has(nomeLayer))
            {
                transacao.Commit();
                return;
            }

            tabela.UpgradeOpen();
            var registro = new LayerTableRecord { Name = nomeLayer };
            tabela.Add(registro);
            transacao.AddNewlyCreatedDBObject(registro, true);
            transacao.Commit();

            resultado.Logs.Add($"Layer '{nomeLayer}' criada.");
        }
        catch (Exception ex)
        {
            try { transacao.Abort(); } catch { }
            resultado.Logs.Add($"Nao foi possivel garantir a layer '{nomeLayer}': {ex.Message}");
        }
    }
}
