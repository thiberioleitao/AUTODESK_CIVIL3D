using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Replica a rotina Dynamo "1.5 - CRIAR REGIOES A PARTIR DE CORREDORES".
///
/// Para cada corredor do desenho e cada baseline dele:
/// (1) coleta estacoes de quebra: cruzamentos com outros alinhamentos +
///     mudancas bruscas de direcao horizontal + mudancas de declividade no
///     perfil de projeto;
/// (2) consolida, filtra por espacamento minimo e remove duplicatas;
/// (3) apaga as regioes existentes e cria uma nova regiao entre cada par
///     consecutivo de estacoes, usando o Assembly escolhido;
/// (4) configura as tres frequencias (horizontal, vertical, offset target)
///     em cada regiao criada.
/// </summary>
public static class CriarRegioesCorredoresService
{
    private const double m_toleranciaEstacao = 0.01;

    public static CriarRegioesCorredoresResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        CriarRegioesCorredoresRequest request)
    {
        var resultado = new CriarRegioesCorredoresResultado();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            ObjectId idAssembly = ObterIdAssemblyPorNome(civilDoc, transacao, request.NomeAssembly);
            if (idAssembly.IsNull)
            {
                resultado.MensagensErro.Add($"Assembly '{request.NomeAssembly}' nao encontrado no desenho.");
                transacao.Commit();
                return resultado;
            }

            var idsOutrosAlinhamentos = new List<ObjectId>();
            foreach (ObjectId id in civilDoc.GetAlignmentIds())
                idsOutrosAlinhamentos.Add(id);

            CorridorCollection colecaoCorredores = civilDoc.CorridorCollection;
            resultado.TotalCorredores = colecaoCorredores.Count;

            if (colecaoCorredores.Count == 0)
            {
                resultado.MensagensErro.Add("Nenhum corredor encontrado no desenho.");
                transacao.Commit();
                return resultado;
            }

            foreach (ObjectId idCorredor in colecaoCorredores)
            {
                if (idCorredor.IsNull || idCorredor.IsErased) continue;

                try
                {
                    if (transacao.GetObject(idCorredor, OpenMode.ForWrite) is not Corridor corredor)
                        continue;

                    ProcessarCorredor(
                        corredor,
                        idAssembly,
                        idsOutrosAlinhamentos,
                        transacao,
                        request,
                        resultado,
                        ed);
                }
                catch (Exception ex)
                {
                    resultado.MensagensErro.Add($"Corredor (id {idCorredor}): erro geral: {ex.Message}");
                }
            }

            transacao.Commit();
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add($"Erro global na execucao: {ex.Message}");
            try { transacao.Abort(); } catch { }
        }

        return resultado;
    }

    private static void ProcessarCorredor(
        Corridor corredor,
        ObjectId idAssembly,
        List<ObjectId> idsOutrosAlinhamentos,
        Transaction transacao,
        CriarRegioesCorredoresRequest request,
        CriarRegioesCorredoresResultado resultado,
        Editor ed)
    {
        ed.WriteMessage($"\n[ZagoCivil3D] Corredor: {corredor.Name}");

        bool corredorAlterado = false;

        foreach (Baseline baseline in corredor.Baselines)
        {
            resultado.TotalBaselines++;

            try
            {
                int regioesCriadas = ProcessarBaseline(
                    baseline,
                    idAssembly,
                    idsOutrosAlinhamentos,
                    transacao,
                    request,
                    resultado,
                    ed);

                if (regioesCriadas > 0)
                {
                    resultado.TotalBaselinesProcessadas++;
                    resultado.TotalRegioesCriadas += regioesCriadas;
                    corredorAlterado = true;
                    resultado.Logs.Add($"'{corredor.Name}' / '{baseline.Name}': {regioesCriadas} regiao(oes) criada(s).");
                }
            }
            catch (Exception ex)
            {
                resultado.MensagensErro.Add(
                    $"'{corredor.Name}' / '{baseline.Name}': {ex.Message}");
            }
        }

        if (corredorAlterado)
        {
            try
            {
                corredor.Rebuild();
            }
            catch (Exception ex)
            {
                resultado.Logs.Add($"'{corredor.Name}': rebuild final falhou ({ex.Message}).");
            }
        }
    }

    /// <summary>
    /// Coleta estacoes de quebra (cruzamentos + direcao + declividade), cria
    /// novas regioes e configura as tres frequencias. Retorna o numero de
    /// regioes efetivamente criadas.
    /// </summary>
    private static int ProcessarBaseline(
        Baseline baseline,
        ObjectId idAssembly,
        List<ObjectId> idsOutrosAlinhamentos,
        Transaction transacao,
        CriarRegioesCorredoresRequest request,
        CriarRegioesCorredoresResultado resultado,
        Editor ed)
    {
        ObjectId idAlinhamento = baseline.AlignmentId;
        if (idAlinhamento.IsNull || idAlinhamento.IsErased)
            return 0;

        if (transacao.GetObject(idAlinhamento, OpenMode.ForRead) is not Alignment alinhamento)
            return 0;

        double estacaoInicial = baseline.StartStation;
        double estacaoFinal = baseline.EndStation;

        if (estacaoFinal - estacaoInicial <= m_toleranciaEstacao)
        {
            resultado.Logs.Add($"'{baseline.Name}': faixa de estacoes invalida, baseline ignorada.");
            return 0;
        }

        // (1) estacoes de cruzamento com outros alinhamentos
        List<double> estacoesCruzamento = ObterEstacoesCruzamentos(
            alinhamento, idAlinhamento, idsOutrosAlinhamentos, transacao, estacaoInicial, estacaoFinal);

        // (2) estacoes com mudanca brusca de direcao (angulo > limite)
        List<double> estacoesDirecao = ObterEstacoesMudancaDirecao(
            alinhamento, estacaoInicial, estacaoFinal, request.LimiteAnguloDirecao);

        // (3) estacoes com mudanca de declividade no perfil de projeto
        List<double> estacoesDeclividade = ObterEstacoesMudancaDeclividade(
            baseline, transacao, request.LimiteMudancaDeclividade);
        estacoesDeclividade = FiltrarPorEspacamentoMinimo(
            estacoesDeclividade, request.EspacamentoMinimoDeclividade);

        // Remove criticas muito proximas de cruzamentos (regra Dynamo)
        List<double> estacoesCriticasFiltradas = RemoverProximasDos(
            estacoesDirecao.Concat(estacoesDeclividade).ToList(),
            estacoesCruzamento,
            request.EspacamentoMinimoCruzamentos);

        // Consolida lista final
        var todas = new List<double> { estacaoInicial, estacaoFinal };
        todas.AddRange(estacoesCruzamento);
        todas.AddRange(estacoesCriticasFiltradas);

        List<double> estacoesQuebra = todas
            .Where(s => s >= estacaoInicial - m_toleranciaEstacao
                         && s <= estacaoFinal + m_toleranciaEstacao)
            .Select(s => Math.Floor(s * 100.0) / 100.0)
            .Distinct()
            .OrderBy(s => s)
            .ToList();

        if (estacoesQuebra.Count < 2)
        {
            resultado.Logs.Add($"'{baseline.Name}': sem estacoes de quebra validas.");
            return 0;
        }

        // Apaga regioes existentes para evitar conflito com as novas
        ApagarRegioesExistentes(baseline, resultado);

        int regioesCriadas = 0;

        for (int i = 0; i < estacoesQuebra.Count - 1; i++)
        {
            double inicio = estacoesQuebra[i];
            double fim = estacoesQuebra[i + 1];

            if (fim - inicio <= m_toleranciaEstacao) continue;

            string nomeRegiao = $"{baseline.Name}.T{i + 1}";

            try
            {
                BaselineRegion regiao = baseline.BaselineRegions.Add(nomeRegiao, idAssembly, inicio, fim);
                ConfigurarFrequencias(regiao, request);
                resultado.NomesRegioesCriadas.Add($"{baseline.Name} -> {nomeRegiao} [{inicio:F2}, {fim:F2}]");
                regioesCriadas++;
            }
            catch (Exception ex)
            {
                resultado.MensagensErro.Add(
                    $"'{baseline.Name}' / '{nomeRegiao}' [{inicio:F2}, {fim:F2}]: {ex.Message}");
            }
        }

        return regioesCriadas;
    }

    // ------------------------------------------------------------------
    // Coleta de estacoes
    // ------------------------------------------------------------------

    /// <summary>
    /// Intersecta a geometria do alinhamento com cada outro alinhamento do
    /// desenho, convertendo cada ponto de cruzamento para estacao no proprio
    /// alinhamento.
    /// </summary>
    private static List<double> ObterEstacoesCruzamentos(
        Alignment alinhamento,
        ObjectId idAlinhamento,
        List<ObjectId> idsOutrosAlinhamentos,
        Transaction transacao,
        double estacaoInicial,
        double estacaoFinal)
    {
        var estacoes = new List<double>();

        foreach (ObjectId idOutro in idsOutrosAlinhamentos)
        {
            if (idOutro.IsNull || idOutro.IsErased) continue;
            if (idOutro == idAlinhamento) continue;

            try
            {
                if (transacao.GetObject(idOutro, OpenMode.ForRead) is not Alignment outro)
                    continue;

                var pontos = new Point3dCollection();
                alinhamento.IntersectWith(
                    outro,
                    Intersect.OnBothOperands,
                    pontos,
                    IntPtr.Zero,
                    IntPtr.Zero);

                foreach (Point3d pt in pontos)
                {
                    try
                    {
                        double sta = 0, off = 0;
                        alinhamento.StationOffset(pt.X, pt.Y, ref sta, ref off);
                        if (sta >= estacaoInicial - m_toleranciaEstacao
                            && sta <= estacaoFinal + m_toleranciaEstacao)
                            estacoes.Add(sta);
                    }
                    catch { }
                }
            }
            catch
            {
                // intersecao pode falhar em alinhamentos atipicos; ignoramos
            }
        }

        return estacoes;
    }

    /// <summary>
    /// Retorna as estacoes dos pontos de intersecao de tangentes (PIs) cuja
    /// variacao de direcao entre segmentos adjacentes supera o limite
    /// informado (graus).
    /// </summary>
    private static List<double> ObterEstacoesMudancaDirecao(
        Alignment alinhamento,
        double estacaoInicial,
        double estacaoFinal,
        double limiteAnguloGraus)
    {
        var estacoes = new List<double>();

        Station[] pis;
        try
        {
            pis = alinhamento.GetStationSet(StationTypes.PIPoint);
        }
        catch
        {
            return estacoes;
        }

        if (pis == null || pis.Length == 0)
            return estacoes;

        double limiteRad = limiteAnguloGraus * Math.PI / 180.0;
        // amostra tangente a +/- 0.5 m do PI para detectar mudanca de direcao
        const double deltaAmostra = 0.5;

        foreach (Station pi in pis)
        {
            double estacaoPI = pi.RawStation;
            double staAntes = estacaoPI - deltaAmostra;
            double staDepois = estacaoPI + deltaAmostra;

            if (staAntes < estacaoInicial || staDepois > estacaoFinal) continue;

            try
            {
                double eastingA = 0, northingA = 0, bearingAntes = 0;
                double eastingD = 0, northingD = 0, bearingDepois = 0;
                alinhamento.PointLocation(staAntes, 0, 0.001,
                    ref eastingA, ref northingA, ref bearingAntes);
                alinhamento.PointLocation(staDepois, 0, 0.001,
                    ref eastingD, ref northingD, ref bearingDepois);

                double delta = NormalizarAnguloRad(bearingDepois - bearingAntes);
                if (Math.Abs(delta) > limiteRad)
                    estacoes.Add(estacaoPI);
            }
            catch
            {
            }
        }

        return estacoes;
    }

    /// <summary>
    /// Retorna as estacoes dos PVIs do perfil de projeto da baseline onde a
    /// mudanca de declividade (|GradeOut - GradeIn|) ultrapassa o limite.
    /// </summary>
    private static List<double> ObterEstacoesMudancaDeclividade(
        Baseline baseline,
        Transaction transacao,
        double limiteMudanca)
    {
        var estacoes = new List<double>();

        try
        {
            ObjectId idPerfil = baseline.ProfileId;
            if (idPerfil.IsNull || idPerfil.IsErased) return estacoes;

            if (transacao.GetObject(idPerfil, OpenMode.ForRead) is not Profile perfil)
                return estacoes;

            foreach (ProfilePVI pvi in perfil.PVIs)
            {
                double mudanca = Math.Abs(pvi.GradeOut - pvi.GradeIn);
                if (mudanca > limiteMudanca)
                    estacoes.Add(pvi.Station);
            }
        }
        catch
        {
        }

        return estacoes;
    }

    // ------------------------------------------------------------------
    // Filtros
    // ------------------------------------------------------------------

    /// <summary>
    /// Preserva a primeira estacao e descarta as seguintes que estejam a menos
    /// de <paramref name="espacamento"/> da ultima mantida. Equivalente ao
    /// filtro do Python no Dynamo.
    /// </summary>
    private static List<double> FiltrarPorEspacamentoMinimo(List<double> estacoes, double espacamento)
    {
        if (estacoes == null || estacoes.Count == 0)
            return new List<double>();

        List<double> ordenadas = estacoes.OrderBy(x => x).ToList();
        var resultado = new List<double> { ordenadas[0] };
        double ultima = ordenadas[0];

        for (int i = 1; i < ordenadas.Count; i++)
        {
            if (ordenadas[i] - ultima >= espacamento)
            {
                resultado.Add(ordenadas[i]);
                ultima = ordenadas[i];
            }
        }

        return resultado;
    }

    /// <summary>
    /// Remove estacoes criticas que estejam a menos de <paramref name="espacamento"/>
    /// de qualquer cruzamento. Evita regioes curtas sobrepondo criticos e PIs.
    /// </summary>
    private static List<double> RemoverProximasDos(
        List<double> criticas,
        List<double> cruzamentos,
        double espacamento)
    {
        if (criticas == null || criticas.Count == 0)
            return new List<double>();

        if (cruzamentos == null || cruzamentos.Count == 0)
            return criticas;

        var resultado = new List<double>();
        foreach (double estacao in criticas)
        {
            bool manter = true;
            foreach (double cruz in cruzamentos)
            {
                if (Math.Abs(estacao - cruz) < espacamento)
                {
                    manter = false;
                    break;
                }
            }
            if (manter) resultado.Add(estacao);
        }
        return resultado;
    }

    // ------------------------------------------------------------------
    // Manipulacao de regioes e frequencias
    // ------------------------------------------------------------------

    private static void ApagarRegioesExistentes(Baseline baseline, CriarRegioesCorredoresResultado resultado)
    {
        try
        {
            // A colecao e alterada conforme removemos, por isso iteramos de tras para frente.
            BaselineRegionCollection regioes = baseline.BaselineRegions;
            for (int i = regioes.Count - 1; i >= 0; i--)
            {
                try
                {
                    regioes.RemoveAt(i);
                }
                catch (Exception ex)
                {
                    resultado.Logs.Add($"'{baseline.Name}': nao apagou regiao {i} ({ex.Message}).");
                }
            }
        }
        catch (Exception ex)
        {
            resultado.Logs.Add($"'{baseline.Name}': falha ao limpar regioes ({ex.Message}).");
        }
    }

    /// <summary>
    /// Configura as tres frequencias (horizontal, vertical, offset target) na
    /// regiao recem criada, via o AppliedAssemblySetting.
    /// </summary>
    private static void ConfigurarFrequencias(BaselineRegion regiao, CriarRegioesCorredoresRequest req)
    {
        try
        {
            AppliedAssemblySetting config = regiao.AppliedAssemblySetting;

            // Horizontal
            config.FrequencyAlongTangents = req.FrequenciaTangente;
            config.CorridorAlongCurvesOption = CorridorAlongCurveOption.CurveAtIncrement;
            config.FrequencyAlongCurves = req.IncrementoCurvaHorizontal;
            config.MODAlongCurves = req.DistanciaMidOrdinate;
            config.FrequencyAlongSpirals = req.FrequenciaEspiral;
            config.AppliedAtHorizontalGeometryPoints = req.HorizontalEmGeometryPoints;
            config.AppliedAtSuperelevationCriticalPoints = req.EmPontosSuperelevacao;

            // Vertical
            config.FrequencyAlongProfileCurves = req.FrequenciaCurvaVertical;
            config.AppliedAtProfileGeometryPoints = req.VerticalEmGeometryPoints;
            config.AppliedAtProfileHighLowPoints = req.VerticalEmHighLowPoints;

            // Offset target
            config.AppliedAtOffsetTargetGeometryPoints = req.OffsetEmGeometryPoints;
            config.AppliedAdjacentToOffsetTargetStartEnd = req.OffsetAdjacenteStartEnd;
            config.TargetCurveOption = CorridorAlongOffsetTargetCurveOption.TargetCurveAtIncrement;
            config.FrequencyAlongTargetCurves = req.IncrementoCurvaOffset;
            config.MODAlongTargetCurves = req.DistanciaMidOrdinateOffset;
        }
        catch (Exception)
        {
            // Falha em frequencias nao deve invalidar a regiao criada.
            // Mantemos o comportamento silencioso como no Dynamo.
        }
    }

    // ------------------------------------------------------------------
    // Utilitarios
    // ------------------------------------------------------------------

    /// <summary>
    /// Normaliza um angulo em radianos para o intervalo (-pi, +pi].
    /// </summary>
    private static double NormalizarAnguloRad(double angulo)
    {
        while (angulo > Math.PI) angulo -= 2 * Math.PI;
        while (angulo < -Math.PI) angulo += 2 * Math.PI;
        return angulo;
    }

    private static ObjectId ObterIdAssemblyPorNome(
        CivilDocument civilDoc,
        Transaction transacao,
        string nome)
    {
        if (string.IsNullOrWhiteSpace(nome)) return ObjectId.Null;

        AssemblyCollection colecao = civilDoc.AssemblyCollection;
        for (int i = 0; i < colecao.Count; i++)
        {
            ObjectId id = colecao[i];
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is Autodesk.Civil.DatabaseServices.Assembly asm
                    && string.Equals(asm.Name, nome, StringComparison.OrdinalIgnoreCase))
                {
                    return id;
                }
            }
            catch { }
        }
        return ObjectId.Null;
    }

    /// <summary>
    /// Lista os nomes de Assemblies no desenho para alimentar o ComboBox da UI.
    /// </summary>
    public static List<string> ObterNomesAssemblies(CivilDocument civilDoc, Database db)
    {
        var nomes = new List<string>();
        using Transaction transacao = db.TransactionManager.StartTransaction();

        AssemblyCollection colecao = civilDoc.AssemblyCollection;
        for (int i = 0; i < colecao.Count; i++)
        {
            ObjectId id = colecao[i];
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is Autodesk.Civil.DatabaseServices.Assembly asm
                    && !string.IsNullOrWhiteSpace(asm.Name))
                {
                    nomes.Add(asm.Name);
                }
            }
            catch { }
        }

        transacao.Commit();
        return nomes.Distinct().OrderBy(x => x).ToList();
    }
}
