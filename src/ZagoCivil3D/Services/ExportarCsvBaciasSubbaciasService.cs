using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Porta em C# da rotina Dynamo
/// "EXPORTAR CSV BACIAS E SUBBACIAS - ID, AREA, TALVEGUE, DECLIVIDADE E ID_JUSANTE".
///
/// Para cada hatch de uma layer "HDR-BACIA Bxx" (uma hatch = uma subbacia) lista:
///   BACIA, SUBBACIA, AREA, L_TALVEGUE, S_TALVEGUE, ID_JUSANTE,
///   VERF_SB, VERF_A, VERF_L, VERF_S
///
/// Reusa a mesma estrategia de ray casting usada em
/// <see cref="CriarCatchmentsDeHatchsService"/> para detectar pontos dentro do
/// contorno da hatch.
/// </summary>
public static class ExportarCsvBaciasSubbaciasService
{
    private const double ToleranciaContornoXY = 1e-6;

    /// <summary>Mensagens de verificacao usadas nas colunas VERF_* do CSV.</summary>
    private const string VerfOk = LinhaCsvBaciasSubbacias.VerfOk;
    private const string VerfSbErro = "REVISAR. POSSIVEL ERRO NA HACHURA. POSSIVEL ERRO NO MTEXT DO ID DA SUBBACIA.";
    private const string VerfAreaErro = "REVISAR. ERRO DE CONVERSAO NA HACHURA, POSSIVELMENTE MAIS DE UMA HACHURA IDENTIFICADA. CONFIRMAR AREA.";
    private const string VerfTalvegueErro = "REVISAR. POSSIVEL ERRO NA HACHURA. POSSIVEL ERRO NO TALVEGUE, CONFIRMAR SE ESTA CONTIDO TOTALMENTE NA HACHURA. CONFIRMAR SE EXISTE TALVEGUE.";

    public static ExportarCsvBaciasSubbaciasResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        ExportarCsvBaciasSubbaciasRequest request)
    {
        var resultado = new ExportarCsvBaciasSubbaciasResultado();

        if (string.IsNullOrWhiteSpace(request.CaminhoCsv))
        {
            resultado.MensagensErro.Add("Caminho do CSV nao informado.");
            return resultado;
        }
        if (string.IsNullOrWhiteSpace(request.PrefixoLayersHatches))
        {
            resultado.MensagensErro.Add("Prefixo das layers das hatches nao informado.");
            return resultado;
        }
        if (string.IsNullOrWhiteSpace(request.NomeLayerMTextsSubbaciasId))
        {
            resultado.MensagensErro.Add("Layer dos MTexts dos IDs das subbacias nao informada.");
            return resultado;
        }

        // Valida diretorio do CSV antes de processar.
        try
        {
            string? diretorio = Path.GetDirectoryName(Path.GetFullPath(request.CaminhoCsv));
            if (!string.IsNullOrWhiteSpace(diretorio) && !Directory.Exists(diretorio))
            {
                resultado.MensagensErro.Add($"Diretorio do CSV nao existe: '{diretorio}'.");
                return resultado;
            }
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add($"Caminho do CSV invalido: {ex.Message}");
            return resultado;
        }

        // Resolve superficie (opcional — sem ela, S_TALVEGUE fica em branco).
        Autodesk.Civil.DatabaseServices.Surface? superficie = null;
        ObjectId idSuperficie = ResolverIdSuperficie(civilDoc, db, request.NomeSuperficie);
        if (idSuperficie.IsNull && !string.IsNullOrWhiteSpace(request.NomeSuperficie))
        {
            resultado.Logs.Add(
                $"Superficie '{request.NomeSuperficie}' nao encontrada. S_TALVEGUE ficara em branco.");
        }

        List<HatchDeSubbacia> hatches;
        List<MTextInfo> mtexts;
        List<TalvegueInfo> talvegues;
        List<BaselineRegionInfo> regioes;

        using (Transaction transacao = db.TransactionManager.StartTransaction())
        {
            if (!idSuperficie.IsNull)
                superficie = transacao.GetObject(idSuperficie, OpenMode.ForRead)
                    as Autodesk.Civil.DatabaseServices.Surface;

            hatches = ColetarHatches(transacao, db, request.PrefixoLayersHatches, resultado);
            mtexts = ColetarMTexts(
                transacao, db,
                request.NomeLayerMTextsSubbaciasId,
                request.PrefixoTextoSubbacia);
            talvegues = ColetarTalvegues(transacao, db, request.NomeLayerTalvegues);
            regioes = ColetarRegioesCorredores(civilDoc, transacao, resultado);

            resultado.TotalHatches = hatches.Count;
            resultado.TotalMTextsIds = mtexts.Count;
            resultado.TotalTalvegues = talvegues.Count;
            resultado.TotalRegioesCorredores = regioes.Count;

            if (hatches.Count == 0)
            {
                resultado.MensagensErro.Add(
                    $"Nenhuma hatch encontrada em layers com prefixo '{request.PrefixoLayersHatches}'.");
                transacao.Commit();
                return resultado;
            }

            // Avisos preventivos: a rotina nao aborta mas a UI fica com muitos
            // VERFs vermelhos se alguma das colecoes vier vazia — melhor
            // sinalizar na barra de status e nos logs.
            if (talvegues.Count == 0 && !string.IsNullOrWhiteSpace(request.NomeLayerTalvegues))
            {
                resultado.MensagensErro.Add(
                    $"Nenhuma polilinha encontrada na layer '{request.NomeLayerTalvegues}'. " +
                    "Todas as subbacias sairao com VERF_L e VERF_S como REVISAR. " +
                    "Confirme o nome exato da layer (inclusive espacos) via comando LAYER.");
            }
            if (mtexts.Count == 0 && !string.IsNullOrWhiteSpace(request.NomeLayerMTextsSubbaciasId))
            {
                resultado.MensagensErro.Add(
                    $"Nenhum MText/DBText encontrado na layer '{request.NomeLayerMTextsSubbaciasId}' " +
                    $"com o prefixo '{request.PrefixoTextoSubbacia}'. Todas as subbacias sairao com VERF_SB como REVISAR.");
            }

            ed.WriteMessage(
                $"\n[ZagoCivil3D] Hatches: {hatches.Count} | MTexts: {mtexts.Count} | " +
                $"Talvegues: {talvegues.Count} | Regioes: {regioes.Count}.");

            // Calcula linhas do CSV ainda dentro da transacao — precisamos da
            // superficie aberta para FindElevationAtXY.
            List<LinhaCsvBaciasSubbacias> linhas = MontarLinhas(
                hatches, mtexts, talvegues, regioes, superficie, request);

            transacao.Commit();

            // Popula Linhas no resultado mesmo quando a escrita falha —
            // a UI aproveita para mostrar preview da tabela na aba de resultados.
            resultado.Linhas.AddRange(linhas);
            resultado.TotalLinhasComAviso = linhas.Count(l => l.TemAviso);

            // Escreve o CSV fora da transacao (IO puro).
            try
            {
                EscreverCsv(request.CaminhoCsv, linhas);
                resultado.TotalLinhasCsv = linhas.Count;
                resultado.CaminhoCsvGerado = Path.GetFullPath(request.CaminhoCsv);
                resultado.Logs.Add($"CSV escrito em '{resultado.CaminhoCsvGerado}'.");
            }
            catch (Exception ex)
            {
                resultado.MensagensErro.Add($"Erro ao gravar CSV: {ex.Message}");
            }
        }

        return resultado;
    }

    // ------------------------------------------------------------------
    // Coleta de entidades
    // ------------------------------------------------------------------

    /// <summary>
    /// Varre o ModelSpace e retorna todas as Hatches cuja layer comece com o
    /// prefixo informado (case-insensitive). O contorno XY de cada hatch e
    /// extraido aqui para evitar reabrir depois.
    /// </summary>
    private static List<HatchDeSubbacia> ColetarHatches(
        Transaction transacao,
        Database db,
        string prefixoLayer,
        ExportarCsvBaciasSubbaciasResultado resultado)
    {
        var lista = new List<HatchDeSubbacia>();

        var bloco = (BlockTableRecord)transacao.GetObject(
            SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

        foreach (ObjectId id in bloco)
        {
            if (id.IsNull || id.IsErased) continue;
            if (transacao.GetObject(id, OpenMode.ForRead) is not Hatch hatch) continue;

            string camada = hatch.Layer ?? string.Empty;
            if (camada.IndexOf(prefixoLayer, StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            (List<Point3d> contorno, bool contornoValido) = ExtrairContornoDaHatch(hatch, resultado);

            lista.Add(new HatchDeSubbacia
            {
                Handle = hatch.Handle.ToString(),
                NomeLayer = camada,
                Area = hatch.Area,
                Contorno = contorno,
                ContornoValido = contornoValido && contorno.Count >= 3,
            });
        }

        return lista;
    }

    /// <summary>
    /// Extrai o contorno XY da primeira loop externa da hatch. Mesma logica do
    /// <see cref="CriarCatchmentsDeHatchsService"/>. O flag retornado indica se
    /// o contorno foi extraido sem excecoes — usado para acender VERF_A.
    /// </summary>
    private static (List<Point3d>, bool) ExtrairContornoDaHatch(
        Hatch hatch,
        ExportarCsvBaciasSubbaciasResultado resultado)
    {
        var pontos = new List<Point3d>();
        bool sucesso = true;

        try
        {
            int numLoops = hatch.NumberOfLoops;
            for (int i = 0; i < numLoops; i++)
            {
                HatchLoop loop = hatch.GetLoopAt(i);

                if ((loop.LoopType & HatchLoopTypes.External) == 0
                    && (loop.LoopType & HatchLoopTypes.Outermost) == 0)
                    continue;

                if (loop.IsPolyline)
                {
                    BulgeVertexCollection verts = loop.Polyline;
                    foreach (BulgeVertex bv in verts)
                        pontos.Add(new Point3d(bv.Vertex.X, bv.Vertex.Y, 0));
                }
                else
                {
                    foreach (Curve2d curva in loop.Curves)
                    {
                        Point2d p0 = curva.StartPoint;
                        Point2d p1 = curva.EndPoint;
                        if (pontos.Count == 0)
                            pontos.Add(new Point3d(p0.X, p0.Y, 0));
                        pontos.Add(new Point3d(p1.X, p1.Y, 0));
                    }
                }

                if (pontos.Count >= 3)
                    break;
            }
        }
        catch (Exception ex)
        {
            resultado.Logs.Add($"Hatch {hatch.Handle}: falha ao extrair contorno ({ex.Message}).");
            sucesso = false;
        }

        // Remove duplicata final se primeiro == ultimo.
        if (pontos.Count >= 2
            && Math.Abs(pontos[0].X - pontos[^1].X) < ToleranciaContornoXY
            && Math.Abs(pontos[0].Y - pontos[^1].Y) < ToleranciaContornoXY)
        {
            pontos.RemoveAt(pontos.Count - 1);
        }

        return (pontos, sucesso);
    }

    /// <summary>
    /// Coleta MTexts/DBTexts da layer informada cujo conteudo contenha o
    /// prefixo indicado (ex.: "SB"). Aceita DBText alem de MText — alguns
    /// projetos usam texto simples.
    /// </summary>
    private static List<MTextInfo> ColetarMTexts(
        Transaction transacao,
        Database db,
        string nomeLayer,
        string prefixoTexto)
    {
        var lista = new List<MTextInfo>();

        var bloco = (BlockTableRecord)transacao.GetObject(
            SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

        foreach (ObjectId id in bloco)
        {
            if (id.IsNull || id.IsErased) continue;

            if (transacao.GetObject(id, OpenMode.ForRead) is MText mt
                && string.Equals(mt.Layer, nomeLayer, StringComparison.OrdinalIgnoreCase))
            {
                string texto = LimparFormatacaoMText(mt.Contents ?? mt.Text ?? string.Empty);
                if (ContemPrefixo(texto, prefixoTexto))
                {
                    lista.Add(new MTextInfo
                    {
                        Conteudo = texto,
                        Posicao = new Point3d(mt.Location.X, mt.Location.Y, 0),
                    });
                }
            }
            else if (transacao.GetObject(id, OpenMode.ForRead) is DBText dbText
                     && string.Equals(dbText.Layer, nomeLayer, StringComparison.OrdinalIgnoreCase))
            {
                string texto = dbText.TextString ?? string.Empty;
                if (ContemPrefixo(texto, prefixoTexto))
                {
                    lista.Add(new MTextInfo
                    {
                        Conteudo = texto.Trim(),
                        Posicao = new Point3d(dbText.Position.X, dbText.Position.Y, 0),
                    });
                }
            }
        }

        return lista;
    }

    private static bool ContemPrefixo(string texto, string prefixo)
    {
        if (string.IsNullOrWhiteSpace(prefixo)) return true;
        return texto.IndexOf(prefixo, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    /// <summary>
    /// Coleta polilinhas da layer informada e guarda vertices + extremos.
    /// </summary>
    private static List<TalvegueInfo> ColetarTalvegues(
        Transaction transacao,
        Database db,
        string nomeLayer)
    {
        var lista = new List<TalvegueInfo>();
        if (string.IsNullOrWhiteSpace(nomeLayer))
            return lista;

        var bloco = (BlockTableRecord)transacao.GetObject(
            SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

        foreach (ObjectId id in bloco)
        {
            if (id.IsNull || id.IsErased) continue;

            var entidade = transacao.GetObject(id, OpenMode.ForRead);
            if (entidade is not Autodesk.AutoCAD.DatabaseServices.Entity ent) continue;
            if (!string.Equals(ent.Layer, nomeLayer, StringComparison.OrdinalIgnoreCase)) continue;

            List<Point3d>? pontos = null;
            double comprimento = 0.0;

            if (entidade is Polyline poli)
            {
                pontos = new List<Point3d>();
                for (int i = 0; i < poli.NumberOfVertices; i++)
                    pontos.Add(poli.GetPoint3dAt(i));
                comprimento = poli.Length;
            }
            else if (entidade is Polyline2d poli2d)
            {
                pontos = new List<Point3d>();
                foreach (ObjectId vid in poli2d)
                {
                    if (transacao.GetObject(vid, OpenMode.ForRead) is Vertex2d v2d)
                        pontos.Add(v2d.Position);
                }
                comprimento = CalcularComprimento(pontos);
            }
            else if (entidade is Polyline3d poli3d)
            {
                pontos = new List<Point3d>();
                foreach (ObjectId vid in poli3d)
                {
                    if (transacao.GetObject(vid, OpenMode.ForRead) is PolylineVertex3d v3d)
                        pontos.Add(v3d.Position);
                }
                comprimento = CalcularComprimento(pontos);
            }
            else if (entidade is Line linha)
            {
                pontos = new List<Point3d> { linha.StartPoint, linha.EndPoint };
                comprimento = linha.Length;
            }

            if (pontos == null || pontos.Count < 2)
                continue;

            lista.Add(new TalvegueInfo
            {
                Inicio = new Point3d(pontos[0].X, pontos[0].Y, 0),
                Fim = new Point3d(pontos[^1].X, pontos[^1].Y, 0),
                Comprimento = comprimento,
            });
        }

        return lista;
    }

    /// <summary>
    /// Soma os comprimentos entre vertices consecutivos (projecao XY+Z).
    /// Usado como fallback quando a entidade nao expoe .Length.
    /// </summary>
    private static double CalcularComprimento(List<Point3d> pontos)
    {
        double total = 0.0;
        for (int i = 1; i < pontos.Count; i++)
            total += pontos[i].DistanceTo(pontos[i - 1]);
        return total;
    }

    /// <summary>
    /// Coleta todas as BaselineRegions de todos os corredores do desenho,
    /// incluindo o alinhamento base e os limites de estacao. Usadas para
    /// localizar o ID_JUSANTE a partir do ponto final do talvegue.
    /// </summary>
    private static List<BaselineRegionInfo> ColetarRegioesCorredores(
        CivilDocument civilDoc,
        Transaction transacao,
        ExportarCsvBaciasSubbaciasResultado resultado)
    {
        var lista = new List<BaselineRegionInfo>();

        CorridorCollection colecao;
        try
        {
            colecao = civilDoc.CorridorCollection;
        }
        catch
        {
            return lista;
        }

        foreach (ObjectId idCorr in colecao)
        {
            if (idCorr.IsNull || idCorr.IsErased) continue;

            try
            {
                if (transacao.GetObject(idCorr, OpenMode.ForRead) is not Corridor corredor)
                    continue;

                foreach (Baseline baseline in corredor.Baselines)
                {
                    ObjectId idAlinh = baseline.AlignmentId;
                    if (idAlinh.IsNull || idAlinh.IsErased) continue;

                    if (transacao.GetObject(idAlinh, OpenMode.ForRead) is not Alignment alinhamento)
                        continue;

                    foreach (BaselineRegion regiao in baseline.BaselineRegions)
                    {
                        lista.Add(new BaselineRegionInfo
                        {
                            Nome = regiao.Name ?? string.Empty,
                            Alinhamento = alinhamento,
                            EstacaoInicial = regiao.StartStation,
                            EstacaoFinal = regiao.EndStation,
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                resultado.Logs.Add($"Corredor {idCorr}: falha ao ler regioes ({ex.Message}).");
            }
        }

        return lista;
    }

    // ------------------------------------------------------------------
    // Montagem das linhas
    // ------------------------------------------------------------------

    private static List<LinhaCsvBaciasSubbacias> MontarLinhas(
        List<HatchDeSubbacia> hatches,
        List<MTextInfo> mtexts,
        List<TalvegueInfo> talvegues,
        List<BaselineRegionInfo> regioes,
        Autodesk.Civil.DatabaseServices.Surface? superficie,
        ExportarCsvBaciasSubbaciasRequest request)
    {
        var linhas = new List<LinhaCsvBaciasSubbacias>();

        foreach (HatchDeSubbacia hatch in hatches)
        {
            var linha = new LinhaCsvBaciasSubbacias
            {
                Bacia = ExtrairBaciaDeLayer(hatch.NomeLayer, request.PrefixoRemoverDoNomeBacia),
                Area = hatch.Area,
                VerfA = hatch.ContornoValido ? VerfOk : VerfAreaErro,
            };

            // SUBBACIA: primeiro MText da lista contido no contorno.
            if (hatch.ContornoValido)
            {
                MTextInfo? mtDentro = mtexts.FirstOrDefault(
                    m => PontoDentroDoPoligono(hatch.Contorno, m.Posicao));
                if (mtDentro != null)
                {
                    linha.Subbacia = mtDentro.Conteudo;
                    linha.VerfSb = VerfOk;
                }
                else
                {
                    linha.VerfSb = VerfSbErro;
                }

                // TALVEGUE: polilinha com inicio E fim dentro do contorno.
                TalvegueInfo? talvegue = talvegues.FirstOrDefault(
                    t => PontoDentroDoPoligono(hatch.Contorno, t.Inicio)
                         && PontoDentroDoPoligono(hatch.Contorno, t.Fim));
                if (talvegue != null)
                {
                    linha.Comprimento = talvegue.Comprimento;
                    linha.VerfL = VerfOk;

                    double? declividade = CalcularDeclividade(talvegue, superficie);
                    if (declividade.HasValue)
                    {
                        linha.Declividade = declividade;
                        linha.VerfS = VerfOk;
                    }
                    else
                    {
                        // Sem superficie (ou erro): nao bloqueia — S fica vazio.
                        linha.VerfS = superficie == null ? VerfOk : VerfTalvegueErro;
                    }

                    linha.IdJusante = LocalizarIdJusante(talvegue.Fim, regioes, request.RaioBuscaJusante);
                }
                else
                {
                    linha.VerfL = VerfTalvegueErro;
                    linha.VerfS = VerfTalvegueErro;
                }
            }
            else
            {
                // Sem contorno valido nao da para testar containment.
                linha.VerfSb = VerfSbErro;
                linha.VerfL = VerfTalvegueErro;
                linha.VerfS = VerfTalvegueErro;
            }

            linhas.Add(linha);
        }

        return linhas;
    }

    /// <summary>
    /// Converte o nome da layer no valor de display da coluna BACIA. Remove
    /// o prefixo informado (ex.: "HDR-") e aplica title case. Ex.: layer
    /// "HDR-BACIA 01" com prefixo "HDR-" vira "Bacia 01".
    /// </summary>
    private static string ExtrairBaciaDeLayer(string nomeLayer, string prefixoRemover)
    {
        if (string.IsNullOrWhiteSpace(nomeLayer)) return string.Empty;
        string resto = nomeLayer.Trim();

        if (!string.IsNullOrWhiteSpace(prefixoRemover)
            && resto.StartsWith(prefixoRemover, StringComparison.OrdinalIgnoreCase))
        {
            resto = resto.Substring(prefixoRemover.Length).Trim();
        }

        return AplicarTitleCase(resto);
    }

    /// <summary>
    /// Title case usando pt-BR ("BACIA 01" -> "Bacia 01"). ToTitleCase exige
    /// texto minusculo para transformar palavras ja em caixa alta.
    /// </summary>
    private static string AplicarTitleCase(string texto)
    {
        if (string.IsNullOrWhiteSpace(texto)) return texto;
        CultureInfo cultura = CultureInfo.GetCultureInfo("pt-BR");
        return cultura.TextInfo.ToTitleCase(texto.ToLower(cultura));
    }

    /// <summary>
    /// Declividade media do talvegue: (zMontante - zJusante) / comprimento.
    /// zMontante = ElevacaoTIN(start); zJusante = ElevacaoTIN(end). Retorna
    /// null quando nao ha superficie ou a consulta falha.
    /// </summary>
    private static double? CalcularDeclividade(
        TalvegueInfo talvegue,
        Autodesk.Civil.DatabaseServices.Surface? superficie)
    {
        if (superficie == null) return null;
        if (talvegue.Comprimento <= 0) return null;

        try
        {
            double zm = superficie.FindElevationAtXY(talvegue.Inicio.X, talvegue.Inicio.Y);
            double zj = superficie.FindElevationAtXY(talvegue.Fim.X, talvegue.Fim.Y);
            return (zm - zj) / talvegue.Comprimento;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Para o ponto final do talvegue, procura a baseline region cujo eixo
    /// passa mais proximo (dentro do raio). O nome da regiao vira ID_JUSANTE.
    /// Espelha a logica do Dynamo que desenhava um circulo de 10 m ao redor
    /// do endpoint e testava intersecao com cada PolyCurve de regiao.
    /// </summary>
    private static string LocalizarIdJusante(
        Point3d pontoFinal,
        List<BaselineRegionInfo> regioes,
        double raio)
    {
        if (regioes.Count == 0) return string.Empty;

        BaselineRegionInfo? melhor = null;
        double melhorDistancia = double.MaxValue;

        foreach (BaselineRegionInfo regiao in regioes)
        {
            double sta = 0, off = 0;
            try
            {
                regiao.Alinhamento.StationOffset(pontoFinal.X, pontoFinal.Y, ref sta, ref off);
            }
            catch
            {
                // Ponto cai fora da projecao do alinhamento — ignora essa regiao.
                continue;
            }

            // StationOffset devolve a estacao mesmo quando o ponto cai fora da
            // regiao — filtramos pelo intervalo [start, end] da regiao.
            if (sta + 1e-6 < regiao.EstacaoInicial || sta - 1e-6 > regiao.EstacaoFinal)
                continue;

            double distancia = Math.Abs(off);
            if (distancia > raio) continue;

            if (distancia < melhorDistancia)
            {
                melhorDistancia = distancia;
                melhor = regiao;
            }
        }

        return melhor?.Nome ?? string.Empty;
    }

    // ------------------------------------------------------------------
    // Escrita do CSV
    // ------------------------------------------------------------------

    private static void EscreverCsv(string caminho, List<LinhaCsvBaciasSubbacias> linhas)
    {
        var sb = new StringBuilder();
        sb.AppendLine("BACIA,SUBBACIA,AREA,L_TALVEGUE,S_TALVEGUE,ID_JUSANTE,VERF_SB,VERF_A,VERF_L,VERF_S");

        foreach (LinhaCsvBaciasSubbacias linha in linhas)
        {
            sb.Append(EscaparCampo(linha.Bacia)).Append(',');
            sb.Append(EscaparCampo(linha.Subbacia)).Append(',');
            sb.Append(FormatarNumero(linha.Area)).Append(',');
            sb.Append(FormatarNumero(linha.Comprimento)).Append(',');
            sb.Append(FormatarNumero(linha.Declividade)).Append(',');
            sb.Append(EscaparCampo(linha.IdJusante)).Append(',');
            sb.Append(EscaparCampo(linha.VerfSb)).Append(',');
            sb.Append(EscaparCampo(linha.VerfA)).Append(',');
            sb.Append(EscaparCampo(linha.VerfL)).Append(',');
            sb.AppendLine(EscaparCampo(linha.VerfS));
        }

        // UTF-8 com BOM — Excel abre os acentos e virgulas corretamente.
        File.WriteAllText(caminho, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
    }

    /// <summary>
    /// Escapa um campo de CSV segundo RFC 4180: se contiver virgula, aspas ou
    /// quebra de linha, envolve em aspas duplas e dobra as aspas internas.
    /// </summary>
    private static string EscaparCampo(string? valor)
    {
        if (string.IsNullOrEmpty(valor)) return string.Empty;
        bool precisaAspas = valor.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0;
        if (!precisaAspas) return valor;
        return "\"" + valor.Replace("\"", "\"\"") + "\"";
    }

    /// <summary>
    /// Numero formatado em cultura invariante (ponto como decimal). Null vira
    /// string vazia — replica o comportamento do Dynamo de deixar a celula
    /// em branco quando o valor nao pode ser calculado.
    /// </summary>
    private static string FormatarNumero(double? valor)
    {
        if (!valor.HasValue) return string.Empty;
        if (double.IsNaN(valor.Value) || double.IsInfinity(valor.Value)) return string.Empty;
        return valor.Value.ToString("R", CultureInfo.InvariantCulture);
    }

    // ------------------------------------------------------------------
    // Utilitarios
    // ------------------------------------------------------------------

    private static ObjectId ResolverIdSuperficie(
        CivilDocument civilDoc,
        Database db,
        string nome)
    {
        if (string.IsNullOrWhiteSpace(nome)) return ObjectId.Null;

        using Transaction transacao = db.TransactionManager.StartTransaction();
        foreach (ObjectId id in civilDoc.GetSurfaceIds())
        {
            if (transacao.GetObject(id, OpenMode.ForRead) is Autodesk.Civil.DatabaseServices.Surface sup
                && string.Equals(sup.Name, nome, StringComparison.OrdinalIgnoreCase))
            {
                transacao.Commit();
                return id;
            }
        }
        transacao.Commit();
        return ObjectId.Null;
    }

    /// <summary>
    /// Ray casting 2D para point-in-polygon. Mesma implementacao usada nos
    /// demais services do plugin — mantem consistencia de criterio.
    /// </summary>
    private static bool PontoDentroDoPoligono(List<Point3d> poligono, Point3d ponto)
    {
        if (poligono.Count < 3) return false;

        double x = ponto.X;
        double y = ponto.Y;
        bool dentro = false;
        int j = poligono.Count - 1;

        for (int i = 0; i < poligono.Count; i++)
        {
            double xi = poligono[i].X, yi = poligono[i].Y;
            double xj = poligono[j].X, yj = poligono[j].Y;

            bool intersecta = ((yi > y) != (yj > y))
                              && (x < ((xj - xi) * (y - yi) / (yj - yi + 1e-18)) + xi);
            if (intersecta)
                dentro = !dentro;

            j = i;
        }

        return dentro;
    }

    /// <summary>
    /// Remove formatacao tipica de MText (cores, fontes, chaves). Mesma
    /// implementacao do <see cref="CriarCatchmentsDeHatchsService"/>.
    /// </summary>
    private static string LimparFormatacaoMText(string bruto)
    {
        if (string.IsNullOrEmpty(bruto)) return string.Empty;

        var sb = new StringBuilder(bruto.Length);
        for (int i = 0; i < bruto.Length; i++)
        {
            char c = bruto[i];
            if (c == '\\' && i + 1 < bruto.Length)
            {
                char pr = bruto[i + 1];
                if (pr == '\\' || pr == '{' || pr == '}')
                {
                    sb.Append(pr);
                    i++;
                    continue;
                }
                int j = i + 1;
                while (j < bruto.Length && bruto[j] != ';' && bruto[j] != ' ')
                    j++;
                if (j < bruto.Length && bruto[j] == ';') j++;
                i = j - 1;
                continue;
            }
            if (c == '{' || c == '}')
                continue;
            sb.Append(c);
        }

        return sb.ToString().Trim();
    }

    // ------------------------------------------------------------------
    // Estruturas internas
    // ------------------------------------------------------------------

    private sealed class HatchDeSubbacia
    {
        public string Handle { get; set; } = string.Empty;
        public string NomeLayer { get; set; } = string.Empty;
        public double Area { get; set; }
        public List<Point3d> Contorno { get; set; } = new();
        public bool ContornoValido { get; set; }
    }

    private sealed class MTextInfo
    {
        public string Conteudo { get; set; } = string.Empty;
        public Point3d Posicao { get; set; }
    }

    private sealed class TalvegueInfo
    {
        public Point3d Inicio { get; set; }
        public Point3d Fim { get; set; }
        public double Comprimento { get; set; }
    }

    private sealed class BaselineRegionInfo
    {
        public string Nome { get; set; } = string.Empty;
        public Alignment Alinhamento { get; set; } = null!;
        public double EstacaoInicial { get; set; }
        public double EstacaoFinal { get; set; }
    }
}
