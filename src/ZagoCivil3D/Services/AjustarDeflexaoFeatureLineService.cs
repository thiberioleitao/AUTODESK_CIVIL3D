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
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Servico de ajuste iterativo de deflexao em feature lines.
///
/// <para>Caracteristicas principais:</para>
/// <list type="bullet">
///   <item>Laco de duas (ou mais) passadas sobre os PI points, alternando o sentido em cada
///         passada para acomodar efeitos de cascata gerados por ajustes anteriores.</item>
///   <item>Passo de ajuste dinamico: cresce enquanto faz progresso, bisecta em overshoot;
///         o valor configurado pelo usuario atua como piso de precisao.</item>
///   <item>Escolha de alvo configuravel: sempre o PI central (PontoCentral), ou entre os
///         tres candidatos (Zi-1, Zi, Zi+1) pelo criterio de menor impacto volumetrico
///         de terraplenagem com bonus de reutilizacao (MultiPonto).</item>
///   <item>Pre-scan e post-scan calculam a deflexao por PI antes e depois do ajuste, para
///         relatorio detalhado (violacoes iniciais/finais, estacoes distintas alteradas,
///         |dZ| medio e maximo, iteracoes por estacao).</item>
/// </list>
///
/// Cada feature line e processada em sua propria transacao para isolar falhas.
/// </summary>
public static class AjustarDeflexaoFeatureLineService
{
    /// <summary>
    /// Tolerancia XY usada para mapear um PI point no indice da colecao AllPoints,
    /// que e o indice aceito por FeatureLine.SetPointElevation(int, double).
    /// </summary>
    private const double ToleranciaMapeamentoXY = 1e-4;

    /// <summary>
    /// Executa o ajuste em todas as feature lines do site selecionado.
    /// </summary>
    public static AjustarDeflexaoFeatureLinesResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        AjustarDeflexaoFeatureLinesRequest request)
    {
        var resultado = new AjustarDeflexaoFeatureLinesResultado();
        var deltasAplicados = new List<double>();

        List<ObjectId> idsFeatureLines = ColetarFeatureLines(civilDoc, db, request, resultado);
        resultado.TotalFeatureLinesConsideradas = idsFeatureLines.Count;

        if (idsFeatureLines.Count == 0)
        {
            resultado.MensagensErro.Add(
                "Nenhuma feature line encontrada com os filtros informados (site/nome).");
            return resultado;
        }

        ed.WriteMessage(
            $"\n[ZagoCivil3D] Ajustando deflexao em {idsFeatureLines.Count} feature line(s) (modo {request.ModoOtimizacao})...");

        foreach (ObjectId idFeatureLine in idsFeatureLines)
            ProcessarFeatureLine(db, idFeatureLine, request, resultado, deltasAplicados);

        if (deltasAplicados.Count > 0)
        {
            resultado.DeltaZMedio = deltasAplicados.Average();
            resultado.DeltaZMaximo = deltasAplicados.Max();
        }

        ed.WriteMessage("\n[ZagoCivil3D] Fluxo de ajuste de deflexao concluido.");
        return resultado;
    }

    /// <summary>
    /// Recupera os ObjectIds das feature lines filtradas por site e nome.
    /// </summary>
    private static List<ObjectId> ColetarFeatureLines(
        CivilDocument civilDoc,
        Database db,
        AjustarDeflexaoFeatureLinesRequest request,
        AjustarDeflexaoFeatureLinesResultado resultado)
    {
        var ids = new List<ObjectId>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        List<Site> sitesAlvo = new();
        if (string.IsNullOrWhiteSpace(request.NomeSite))
        {
            foreach (ObjectId idSite in civilDoc.GetSiteIds())
            {
                if (transacao.GetObject(idSite, OpenMode.ForRead) is Site site)
                    sitesAlvo.Add(site);
            }
        }
        else
        {
            Site? site = ObterSitePorNome(civilDoc, transacao, request.NomeSite);
            if (site == null)
            {
                resultado.MensagensErro.Add($"Site '{request.NomeSite}' nao encontrado.");
                transacao.Commit();
                return ids;
            }
            sitesAlvo.Add(site);
        }

        string filtro = request.FiltroNome?.Trim() ?? string.Empty;

        foreach (Site site in sitesAlvo)
        {
            foreach (ObjectId idFeatureLine in site.GetFeatureLineIds())
            {
                if (transacao.GetObject(idFeatureLine, OpenMode.ForRead) is not FeatureLine featureLine)
                    continue;

                if (filtro.Length > 0
                    && (featureLine.Name == null
                        || featureLine.Name.IndexOf(filtro, StringComparison.OrdinalIgnoreCase) < 0))
                {
                    continue;
                }

                ids.Add(idFeatureLine);
            }
        }

        transacao.Commit();
        return ids;
    }

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
    /// Processa uma unica feature line em sua propria transacao de escrita.
    /// Acumula estatisticas (violacoes, iteracoes, deltas) diretamente no resultado agregado.
    /// </summary>
    private static void ProcessarFeatureLine(
        Database db,
        ObjectId idFeatureLine,
        AjustarDeflexaoFeatureLinesRequest request,
        AjustarDeflexaoFeatureLinesResultado resultado,
        List<double> deltasAplicados)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        if (transacao.GetObject(idFeatureLine, OpenMode.ForWrite) is not FeatureLine featureLine)
        {
            resultado.MensagensErro.Add(
                $"Nao foi possivel abrir a feature line {idFeatureLine.Handle} para escrita.");
            return;
        }

        string nomeFeatureLine = string.IsNullOrWhiteSpace(featureLine.Name)
            ? $"#{idFeatureLine.Handle}"
            : featureLine.Name;

        try
        {
            EstatisticaFeatureLine estat = AjustarDeflexaoFeatureLine(
                featureLine, nomeFeatureLine, request, resultado.Logs, deltasAplicados);

            resultado.TotalViolacoesIniciais += estat.ViolacoesIniciais;
            resultado.TotalViolacoesFinais += estat.ViolacoesFinais;
            resultado.TotalViolacoesFinaisInternas += estat.ViolacoesFinaisInternas;
            resultado.TotalViolacoesFinaisFronteira += estat.ViolacoesFinaisFronteira;
            resultado.TotalEstacoesAlteradas += estat.EstacoesAlteradas.Count;
            resultado.TotalIteracoesAplicadas += estat.IteracoesAplicadas;
            resultado.TotalPontosDentroDoLimite += estat.PontosDentroLimiteInicial;
            resultado.TotalPontosNaoConvergidos += estat.PontosNaoConvergidos;

            if (estat.EstacoesAlteradas.Count > 0)
            {
                resultado.TotalFeatureLinesAjustadas++;
                resultado.NomesAjustadas.Add(nomeFeatureLine);
            }

            // Adiciona os pontos individuais e o resumo da FL no relatorio agregado.
            resultado.RelatorioPontos.AddRange(estat.PontosRelatorio);
            resultado.RelatorioFeatureLines.Add(ConsolidarResumoFL(nomeFeatureLine, estat));

            resultado.Logs.Add(
                $"'{nomeFeatureLine}': violacoes iniciais={estat.ViolacoesIniciais}, finais={estat.ViolacoesFinais} (internas={estat.ViolacoesFinaisInternas}, fronteira={estat.ViolacoesFinaisFronteira}), estacoes alteradas={estat.EstacoesAlteradas.Count}, iteracoes={estat.IteracoesAplicadas}.");

            transacao.Commit();
        }
        catch (System.Exception ex)
        {
            resultado.MensagensErro.Add(
                $"'{nomeFeatureLine}': erro no ajuste de deflexao - {ex.Message}");

            // Registra a FL no relatorio mesmo com erro, para o usuario enxergar a falha na tabela.
            resultado.RelatorioFeatureLines.Add(new RelatorioFeatureLine
            {
                Nome = nomeFeatureLine,
                Status = StatusRelatorioFL.Erro,
            });
        }
    }

    /// <summary>
    /// Agrega os dados per-PI em um unico resumo para a tabela principal.
    /// </summary>
    private static RelatorioFeatureLine ConsolidarResumoFL(
        string nomeFeatureLine,
        EstatisticaFeatureLine estat)
    {
        var deltas = estat.PontosRelatorio
            .Where(p => p.DeltaZAcumulado > 0)
            .Select(p => p.DeltaZAcumulado)
            .ToList();

        StatusRelatorioFL status;
        if (estat.ViolacoesIniciais == 0 && estat.ViolacoesFinais == 0)
            status = StatusRelatorioFL.SemViolacoes;
        else if (estat.ViolacoesFinaisInternas > 0)
            status = StatusRelatorioFL.ViolacoesInternasRemanescentes;
        else if (estat.ViolacoesFinaisFronteira > 0)
            status = StatusRelatorioFL.ViolacoesApenasEmFronteira;
        else
            status = StatusRelatorioFL.AjustadoComSucesso;

        return new RelatorioFeatureLine
        {
            Nome = nomeFeatureLine,
            TotalPIAvaliados = estat.PontosRelatorio.Count,
            ViolacoesIniciais = estat.ViolacoesIniciais,
            ViolacoesFinais = estat.ViolacoesFinais,
            ViolacoesFinaisInternas = estat.ViolacoesFinaisInternas,
            ViolacoesFinaisFronteira = estat.ViolacoesFinaisFronteira,
            EstacoesAlteradas = estat.EstacoesAlteradas.Count,
            Iteracoes = estat.IteracoesAplicadas,
            DeltaZMaximo = deltas.Count > 0 ? deltas.Max() : 0,
            DeltaZMedio = deltas.Count > 0 ? deltas.Average() : 0,
            Status = status,
        };
    }

    /// <summary>
    /// Laco principal por feature line. Faz pre-scan (deflexao inicial por PI), executa as
    /// passadas registrando alvos e iteracoes, depois faz post-scan para consolidar a deflexao
    /// final por PI e montar o relatorio detalhado.
    /// </summary>
    private static EstatisticaFeatureLine AjustarDeflexaoFeatureLine(
        FeatureLine featureLine,
        string nomeFeatureLine,
        AjustarDeflexaoFeatureLinesRequest request,
        ICollection<string> logs,
        List<double> deltasAplicados)
    {
        var estatistica = new EstatisticaFeatureLine();

        Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);
        int quantidadePontos = pontosPi.Count;

        if (quantidadePontos < 3)
        {
            logs.Add($"'{nomeFeatureLine}': possui menos de 3 PI points, ajuste ignorado.");
            return estatistica;
        }

        int ultimoIndice = Math.Min(request.IndiceMaximo, quantidadePontos - 2);
        if (ultimoIndice < 1)
        {
            logs.Add($"'{nomeFeatureLine}': indice maximo menor que 1, nada a processar.");
            return estatistica;
        }

        // Pre-scan: calcula deflexao inicial por PI e conta violacoes.
        var estados = new Dictionary<int, EstadoPI>();
        int ultimoPIAvaliavel = quantidadePontos - 2;
        for (int i = 1; i <= ultimoIndice; i++)
        {
            double? defl = CalcularDeflexao(pontosPi[i - 1], pontosPi[i], pontosPi[i + 1]);
            estados[i] = new EstadoPI
            {
                DeflexaoInicial = defl,
                EhFronteira = (i == 1) || (i == ultimoPIAvaliavel),
            };
            if (defl.HasValue && Math.Abs(defl.Value) > request.DeflexaoLimite)
                estatistica.ViolacoesIniciais++;
        }

        int passadasTotais = Math.Max(1, request.QuantidadePassadas);

        for (int passada = 1; passada <= passadasTotais; passada++)
        {
            IEnumerable<int> indices = passada % 2 == 1
                ? Enumerable.Range(1, ultimoIndice)
                : Enumerable.Range(1, ultimoIndice).Reverse();

            foreach (int indicePi in indices)
            {
                ResultadoPonto resultadoPonto = ProcessarPontoDeflexao(
                    featureLine,
                    nomeFeatureLine,
                    indicePi,
                    request,
                    estatistica,
                    estados,
                    deltasAplicados,
                    logs);

                if (resultadoPonto == ResultadoPonto.DentroDoLimite)
                    estatistica.PontosDentroLimiteInicial++;
                else if (resultadoPonto == ResultadoPonto.NaoConvergiu)
                    estatistica.PontosNaoConvergidos++;
            }
        }

        // Post-scan: recalcula deflexao final por PI, coleta posicao XYZ e consolida o relatorio.
        Point3dCollection pontosFinais = featureLine.GetPoints(FeatureLinePointType.PIPoint);
        for (int i = 1; i <= ultimoIndice; i++)
        {
            double? defl = CalcularDeflexao(pontosFinais[i - 1], pontosFinais[i], pontosFinais[i + 1]);
            estados[i].DeflexaoFinal = defl;
            estados[i].PosicaoFinal = pontosFinais[i];
        }

        // Monta o relatorio por PI e os contadores de violacoes finais.
        foreach (int i in estados.Keys.OrderBy(k => k))
        {
            EstadoPI est = estados[i];

            bool violaFinal = est.DeflexaoFinal.HasValue
                              && Math.Abs(est.DeflexaoFinal.Value) > request.DeflexaoLimite;
            bool violavaInicial = est.DeflexaoInicial.HasValue
                                  && Math.Abs(est.DeflexaoInicial.Value) > request.DeflexaoLimite;

            if (violaFinal)
            {
                estatistica.ViolacoesFinais++;
                if (est.EhFronteira)
                    estatistica.ViolacoesFinaisFronteira++;
                else
                    estatistica.ViolacoesFinaisInternas++;
            }

            StatusAjustePI status;
            if (!est.DeflexaoFinal.HasValue || !est.DeflexaoInicial.HasValue)
                status = StatusAjustePI.FalhaGeometrica;
            else if (violaFinal)
                status = StatusAjustePI.ViolacaoRemanescente;
            else if (violavaInicial)
                status = StatusAjustePI.AjustadoComSucesso;
            else
                status = StatusAjustePI.JaDentroDoLimite;

            estatistica.PontosRelatorio.Add(new RelatorioPontoPI
            {
                NomeFeatureLine = nomeFeatureLine,
                IndicePI = i,
                TotalPIs = quantidadePontos,
                DeflexaoInicial = est.DeflexaoInicial,
                DeflexaoFinal = est.DeflexaoFinal,
                DeltaZAcumulado = est.DeltaZAcumulado,
                VezesAlvo = est.VezesAlvo,
                IteracoesAcumuladas = est.IteracoesAcumuladas,
                Status = status,
                EhFronteiraInterna = est.EhFronteira,
                UltimoAlvoEscolhido = est.UltimoAlvoEscolhido,
                PosicaoX = est.PosicaoFinal.X,
                PosicaoY = est.PosicaoFinal.Y,
                PosicaoZ = est.PosicaoFinal.Z,
            });
        }

        return estatistica;
    }

    /// <summary>
    /// Calcula a deflexao no PI de indice <paramref name="indicePi"/> dentro da colecao
    /// de PI points, cuidando de indices fora de faixa (extremos sem dois vizinhos).
    /// </summary>
    private static double? CalcularDeflexaoEmIndice(Point3dCollection pontos, int indicePi)
    {
        if (indicePi < 1 || indicePi + 1 >= pontos.Count)
            return null;
        return CalcularDeflexao(pontos[indicePi - 1], pontos[indicePi], pontos[indicePi + 1]);
    }

    /// <summary>
    /// Calcula D = grade_anterior - grade_posterior em um PI point. Retorna null quando
    /// algum segmento adjacente tem comprimento 2D zero (nao avaliavel).
    /// </summary>
    private static double? CalcularDeflexao(Point3d pAnterior, Point3d pAtual, Point3d pPosterior)
    {
        double Li = Distancia2d(pAnterior, pAtual);
        double Lp1 = Distancia2d(pAtual, pPosterior);
        if (Li <= ToleranciaMapeamentoXY || Lp1 <= ToleranciaMapeamentoXY)
            return null;
        return (pAtual.Z - pAnterior.Z) / Li - (pPosterior.Z - pAtual.Z) / Lp1;
    }

    /// <summary>
    /// Avalia um PI point e, se estiver fora do limite, aplica o ajuste escolhido pelo modo
    /// (tradicional = sempre Zi; inteligente = melhor candidato entre Zi-1/Zi/Zi+1).
    /// </summary>
    private static ResultadoPonto ProcessarPontoDeflexao(
        FeatureLine featureLine,
        string nomeFeatureLine,
        int indicePi,
        AjustarDeflexaoFeatureLinesRequest request,
        EstatisticaFeatureLine estatistica,
        Dictionary<int, EstadoPI> estados,
        List<double> deltasAplicados,
        ICollection<string> logs)
    {
        Point3dCollection pontosPi = featureLine.GetPoints(FeatureLinePointType.PIPoint);

        Point3d pontoAnterior = pontosPi[indicePi - 1];
        Point3d pontoAtual = pontosPi[indicePi];
        Point3d pontoPosterior = pontosPi[indicePi + 1];

        double Li = Distancia2d(pontoAnterior, pontoAtual);
        double Lp1 = Distancia2d(pontoAtual, pontoPosterior);

        if (Li <= ToleranciaMapeamentoXY || Lp1 <= ToleranciaMapeamentoXY)
        {
            logs.Add($"'{nomeFeatureLine}': PI {indicePi} tem segmento adjacente com comprimento 2D nulo - ignorado.");
            return ResultadoPonto.FalhaGeometrica;
        }

        double deflexao = (pontoAtual.Z - pontoAnterior.Z) / Li
                          - (pontoPosterior.Z - pontoAtual.Z) / Lp1;

        if (Math.Abs(deflexao) <= request.DeflexaoLimite)
            return ResultadoPonto.DentroDoLimite;

        AlvoAjuste alvo = SelecionarAlvoAjuste(
            request.ModoOtimizacao,
            indicePi,
            pontoAnterior,
            pontoAtual,
            pontoPosterior,
            Li,
            Lp1,
            deflexao,
            request.DeflexaoLimite,
            estatistica.EstacoesAlteradas,
            request.FatorReutilizacaoEstacao,
            pontosPi);

        // Registra no estado do PI VIOLADOR qual alvo foi escolhido. Deixa visivel na
        // tabela quando um PI em violacao teve seu ajuste delegado a um vizinho.
        if (estados.TryGetValue(indicePi, out EstadoPI? estadoViolador) && estadoViolador != null)
            estadoViolador.UltimoAlvoEscolhido = alvo.IndicePi;

        ResultadoIteracao iteracao = IterarComPassoDinamico(
            alvo,
            deflexao,
            request.DeflexaoLimite,
            request.PassoIncrementalAjusteCota,
            request.PassoMaximoDinamico,
            request.NumeroMaximoIteracoesPorPonto);

        Point3d pontoAlvo = alvo.IndicePi == indicePi - 1
            ? pontoAnterior
            : alvo.IndicePi == indicePi + 1 ? pontoPosterior : pontoAtual;

        ResultadoAjustePonto aplicacao = AtribuirElevacaoPiPoint(featureLine, pontoAlvo, iteracao.Z);
        if (!aplicacao.Sucesso)
        {
            logs.Add($"'{nomeFeatureLine}': PI {alvo.IndicePi} - {aplicacao.Mensagem}");
            return ResultadoPonto.FalhaGeometrica;
        }

        estatistica.EstacoesAlteradas.Add(alvo.IndicePi);
        estatistica.IteracoesAplicadas += iteracao.Iteracoes;

        double deltaZ = Math.Abs(iteracao.Z - alvo.ElevacaoInicial);
        if (deltaZ > 0)
            deltasAplicados.Add(deltaZ);

        // Registra no estado do PI alvo (quando dentro da faixa avaliada).
        if (estados.TryGetValue(alvo.IndicePi, out EstadoPI? estadoAlvo) && estadoAlvo != null)
        {
            estadoAlvo.VezesAlvo++;
            estadoAlvo.DeltaZAcumulado += deltaZ;
            estadoAlvo.IteracoesAcumuladas += iteracao.Iteracoes;
        }

        if (!iteracao.Convergiu)
        {
            logs.Add(
                $"'{nomeFeatureLine}': PI {alvo.IndicePi} nao convergiu apos {iteracao.Iteracoes} iteracoes (|defl| inicial = {Math.Abs(deflexao).ToString("F6", CultureInfo.InvariantCulture)}).");
            return ResultadoPonto.NaoConvergiu;
        }

        return ResultadoPonto.Ajustado;
    }

    /// <summary>
    /// Seleciona qual PI sera ajustado para acomodar a violacao detectada. No modo
    /// <see cref="ModoOtimizacaoDeflexao.PontoCentral"/>, o alvo e sempre o PI central (i).
    /// No modo <see cref="ModoOtimizacaoDeflexao.MultiPonto"/>, os tres candidatos (i-1, i, i+1)
    /// recebem um custo volumetrico ponderado com bonus de reutilizacao; o de menor custo vence.
    /// </summary>
    private static AlvoAjuste SelecionarAlvoAjuste(
        ModoOtimizacaoDeflexao modo,
        int indicePi,
        Point3d pAnt,
        Point3d pAtual,
        Point3d pPost,
        double Li,
        double Lp1,
        double deflexao,
        double deflexaoLimite,
        ISet<int> estacoesAlteradas,
        double fatorReuso,
        Point3dCollection pontosPi)
    {
        // Sempre calcula o candidato "mover Zi" como fallback/tradicional.
        // dD/dZi = 1/Li + 1/Lp1 > 0  =>  SinalDerivada = +1.
        AlvoAjuste candidatoCentro = new(
            indicePi,
            pAtual.Z,
            z => (z - pAnt.Z) / Li - (pPost.Z - z) / Lp1,
            SinalDerivada: +1);

        if (modo == ModoOtimizacaoDeflexao.PontoCentral)
            return candidatoCentro;

        double excesso = Math.Abs(deflexao) - deflexaoLimite;
        // excesso > 0 aqui (so chegamos neste metodo em violacao)

        // Custo proxy de volume de terraplenagem para cada candidato:
        //   a (mover Zi):   |dZa| = excesso * Li*Lp1/(Li+Lp1); influencia ≈ (Li+Lp1)/2
        //                   Va    = excesso * Li*Lp1/2
        //   b (mover Zi-1): |dZb| = excesso * Li;             influencia local = Li
        //                   Vb    = excesso * Li^2
        //   c (mover Zi+1): |dZc| = excesso * Lp1;            influencia local = Lp1
        //                   Vc    = excesso * Lp1^2
        double custoCentro = excesso * Li * Lp1 * 0.5;
        double custoAnterior = excesso * Li * Li;
        double custoPosterior = excesso * Lp1 * Lp1;

        // Bonus de reutilizacao: estacoes ja alteradas tem custo reduzido para concentrar alteracoes.
        //
        // Salvaguarda anti-oscilacao preditiva: estima o impacto que a reutilizacao do candidato
        // teria na deflexao do proprio PI do candidato, e desabilita o bonus se a reutilizacao
        // "quebraria" a restricao que o candidato ja satisfaz.
        //
        // Para mover o vizinho (Zi-1 ou Zi+1) resolvendo uma violacao de magnitude <excesso>
        // no PI sob ajuste, o deslocamento necessario e da ordem de (excesso × L_adjacente).
        // A deflexao do proprio PI do vizinho e afetada por aproximadamente 2 × excesso
        // (a derivada combinada no vizinho e ~2/L_adjacente × deslocamento).
        //
        // Portanto, se |D_candidato_atual| + 2 × excesso > limite, aplicar o bonus levaria o
        // vizinho para fora do limite na proxima iteracao — pagamos o custo cheio para evitar
        // que ele seja escolhido, dirigindo o ajuste para outro candidato.
        double fator = Math.Max(0.01, Math.Min(1.0, fatorReuso));

        double FatorParaCandidato(int indiceEstacao)
        {
            if (!estacoesAlteradas.Contains(indiceEstacao))
                return 1.0; // estacao nunca foi alvo — nao se aplica bonus

            // O proprio PI sob ajuste (centro) pode ser reutilizado livremente: o iterador
            // leva D do proprio PI de volta ao limite, sem quebrar restricao de terceiros.
            if (indiceEstacao == indicePi)
                return fator;

            double? deflexaoCandidata = CalcularDeflexaoEmIndice(pontosPi, indiceEstacao);
            if (!deflexaoCandidata.HasValue)
                return fator;

            double impactoEstimado = 2.0 * excesso;
            double deflexaoProjetada = Math.Abs(deflexaoCandidata.Value) + impactoEstimado;
            if (deflexaoProjetada > deflexaoLimite)
                return 1.0; // reutilizacao quebraria o proprio PI do vizinho — sem bonus

            return fator;
        }

        custoCentro *= FatorParaCandidato(indicePi);
        custoAnterior *= FatorParaCandidato(indicePi - 1);
        custoPosterior *= FatorParaCandidato(indicePi + 1);

        double custoMinimo = Math.Min(custoCentro, Math.Min(custoAnterior, custoPosterior));

        if (custoMinimo == custoAnterior && custoAnterior < custoCentro)
        {
            // dD/dZi-1 = -1/Li < 0  =>  SinalDerivada = -1.
            return new AlvoAjuste(
                indicePi - 1,
                pAnt.Z,
                z => (pAtual.Z - z) / Li - (pPost.Z - pAtual.Z) / Lp1,
                SinalDerivada: -1);
        }

        if (custoMinimo == custoPosterior && custoPosterior < custoCentro)
        {
            // dD/dZi+1 = -1/Lp1 < 0  =>  SinalDerivada = -1.
            return new AlvoAjuste(
                indicePi + 1,
                pPost.Z,
                z => (pAtual.Z - pAnt.Z) / Li - (z - pAtual.Z) / Lp1,
                SinalDerivada: -1);
        }

        return candidatoCentro;
    }

    /// <summary>
    /// Ajusta iterativamente a elevacao do alvo ate que |deflexao| &lt;= limite.
    /// Usa passo dinamico: parte do passo minimo, cresce ao fazer progresso, e bisecta
    /// ao detectar overshoot (mudanca de sinal da deflexao). O passo jamais ultrapassa
    /// <paramref name="passoMaximo"/>.
    /// </summary>
    private static ResultadoIteracao IterarComPassoDinamico(
        AlvoAjuste alvo,
        double deflexaoInicial,
        double deflexaoLimite,
        double passoMinimo,
        double passoMaximo,
        int maxIteracoes)
    {
        double z = alvo.ElevacaoInicial;
        double deflexaoAtual = deflexaoInicial;

        if (Math.Abs(deflexaoAtual) <= deflexaoLimite)
            return new ResultadoIteracao(z, Convergiu: true, Iteracoes: 0);

        // Garantias minimas dos parametros.
        double passoMin = Math.Max(1e-6, passoMinimo);
        double passoMax = Math.Max(passoMin, passoMaximo);

        double passo = passoMin;
        // Para reduzir D (quando D > 0), queremos dZ tal que SinalDerivada * dZ < 0,
        // ou seja, dZ com sinal oposto ao SinalDerivada. Quando D < 0, inverte.
        // Formula geral:  direcao = (D > 0 ? -1 : +1) * SinalDerivada.
        int direcao = (deflexaoAtual > 0 ? -1 : 1) * alvo.SinalDerivada;

        int limite = Math.Max(1, maxIteracoes);
        int iteracoes = 0;

        while (iteracoes < limite)
        {
            iteracoes++;

            double zTeste = z + direcao * passo;
            double deflexaoTeste = alvo.CalcularDeflexao(zTeste);

            if (Math.Abs(deflexaoTeste) <= deflexaoLimite)
                return new ResultadoIteracao(zTeste, Convergiu: true, Iteracoes: iteracoes);

            bool overshoot = Math.Sign(deflexaoTeste) != Math.Sign(deflexaoAtual)
                             && deflexaoTeste != 0
                             && deflexaoAtual != 0;

            if (overshoot)
            {
                // Passou do zero. Aceita zTeste, inverte direcao e reduz passo a metade.
                z = zTeste;
                deflexaoAtual = deflexaoTeste;
                direcao = -direcao;
                passo *= 0.5;

                // Nao deixamos o passo ficar menor que 10% do piso para nao congelar
                // o ajuste quando a deflexao remanescente for < limite mas > precisao.
                if (passo < passoMin * 0.1)
                    return new ResultadoIteracao(z, Convergiu: Math.Abs(deflexaoAtual) <= deflexaoLimite, Iteracoes: iteracoes);
            }
            else if (Math.Abs(deflexaoTeste) < Math.Abs(deflexaoAtual))
            {
                // Progresso na direcao correta: aceita zTeste e tenta crescer o passo.
                z = zTeste;
                deflexaoAtual = deflexaoTeste;
                passo = Math.Min(passoMax, passo * 1.5);
            }
            else
            {
                // Sem progresso (improvavel em D(z) linear, mas cobre edge cases de SetPointElevation):
                // reduz o passo para refinar a tentativa.
                passo *= 0.5;
                if (passo < passoMin * 0.1)
                    return new ResultadoIteracao(z, Convergiu: false, Iteracoes: iteracoes);
            }
        }

        return new ResultadoIteracao(z, Convergiu: Math.Abs(deflexaoAtual) <= deflexaoLimite, Iteracoes: iteracoes);
    }

    /// <summary>
    /// Mapeia o ponto PI para o indice correspondente em AllPoints (unico indice aceito
    /// por FeatureLine.SetPointElevation) e aplica a nova elevacao.
    /// </summary>
    private static ResultadoAjustePonto AtribuirElevacaoPiPoint(
        FeatureLine featureLine,
        Point3d pontoPi,
        double novaElevacao)
    {
        Point3dCollection todosOsPontos = featureLine.GetPoints(FeatureLinePointType.AllPoints);

        for (int indice = 0; indice < todosOsPontos.Count; indice++)
        {
            Point3d ponto = todosOsPontos[indice];
            double deltaX = ponto.X - pontoPi.X;
            double deltaY = ponto.Y - pontoPi.Y;
            if (Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY)) > ToleranciaMapeamentoXY)
                continue;

            try
            {
                featureLine.SetPointElevation(indice, novaElevacao);
                return ResultadoAjustePonto.ComSucesso();
            }
            catch (System.Exception ex)
            {
                return ResultadoAjustePonto.Falha(
                    $"falha ao aplicar elevacao em AllPoints[{indice}]: {ex.Message}");
            }
        }

        return ResultadoAjustePonto.Falha(
            $"falha de mapeamento PI -> AllPoints (tolerancia {ToleranciaMapeamentoXY.ToString("G", CultureInfo.InvariantCulture)}).");
    }

    /// <summary>
    /// Calcula a distancia 2D entre dois pontos, ignorando Z.
    /// </summary>
    private static double Distancia2d(Point3d a, Point3d b)
    {
        double deltaX = b.X - a.X;
        double deltaY = b.Y - a.Y;
        return Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
    }

    /// <summary>
    /// Representa o alvo do ajuste: qual PI mexer, sua elevacao original, como recalcular
    /// a deflexao em funcao da nova elevacao aplicada, e o sinal da derivada dD/dZ.
    ///
    /// <para>
    /// <b>Por que precisamos do SinalDerivada</b>: a deflexao D(z) no PI i depende linearmente
    /// do Z de qualquer um dos tres pontos envolvidos (i-1, i, i+1), mas com sinais diferentes:
    /// </para>
    /// <list type="bullet">
    /// <item>dD/dZi   = 1/Li + 1/Lp1  &gt; 0  (mover Zi na mesma direcao de D aumenta |D|)</item>
    /// <item>dD/dZi-1 = -1/Li         &lt; 0  (mover Zi-1 na mesma direcao de D DIMINUI |D|)</item>
    /// <item>dD/dZi+1 = -1/Lp1        &lt; 0  (idem para Zi+1)</item>
    /// </list>
    /// <para>
    /// O iterador precisa saber o sinal para escolher a direcao inicial correta do passo;
    /// caso contrario, quando o alvo e um vizinho, o passo anda para o lado errado e o
    /// iterador sai como "nao convergiu" apos bisectar ate o piso.
    /// </para>
    /// </summary>
    private sealed record AlvoAjuste(
        int IndicePi,
        double ElevacaoInicial,
        Func<double, double> CalcularDeflexao,
        int SinalDerivada);

    /// <summary>
    /// Resultado da iteracao local para um unico alvo.
    /// </summary>
    private readonly record struct ResultadoIteracao(double Z, bool Convergiu, int Iteracoes);

    /// <summary>
    /// Categorias de resultado de cada PI point avaliado.
    /// </summary>
    private enum ResultadoPonto
    {
        DentroDoLimite,
        Ajustado,
        NaoConvergiu,
        FalhaGeometrica
    }

    /// <summary>
    /// Estatisticas internas por feature line. "EstacoesAlteradas" e um HashSet dos indices
    /// de PI efetivamente alterados (conta cada estacao apenas uma vez, mesmo ajustada
    /// em varias iteracoes ou passadas).
    /// </summary>
    private sealed class EstatisticaFeatureLine
    {
        public int ViolacoesIniciais;
        public int ViolacoesFinais;
        public int ViolacoesFinaisInternas;
        public int ViolacoesFinaisFronteira;
        public int PontosDentroLimiteInicial;
        public int PontosNaoConvergidos;
        public int IteracoesAplicadas;
        public HashSet<int> EstacoesAlteradas { get; } = new();
        public List<RelatorioPontoPI> PontosRelatorio { get; } = new();
    }

    /// <summary>
    /// Estado mutavel de um PI ao longo da execucao: deflexao antes/depois, quantas vezes
    /// foi alvo, soma dos |dZ| aplicados na estacao, e o ultimo alvo escolhido quando
    /// este PI violava (pode ser o proprio indice, ou um vizinho no modo MultiPonto).
    /// </summary>
    private sealed class EstadoPI
    {
        public double? DeflexaoInicial { get; set; }
        public double? DeflexaoFinal { get; set; }
        public double DeltaZAcumulado { get; set; }
        public int VezesAlvo { get; set; }
        public bool EhFronteira { get; set; }

        /// <summary>
        /// Indice do PI que foi movido quando este PI estava violando.
        /// Null quando este PI nunca violou durante as passadas.
        /// Pode ser igual ao proprio indice (ajuste direto) ou um vizinho (modo MultiPonto).
        /// </summary>
        public int? UltimoAlvoEscolhido { get; set; }

        /// <summary>
        /// Soma das iteracoes internas quando este PI foi o alvo do iterador.
        /// Custo de convergencia desta estacao.
        /// </summary>
        public int IteracoesAcumuladas { get; set; }

        /// <summary>Posicao XYZ apos todas as passadas (coletada no post-scan).</summary>
        public Point3d PosicaoFinal { get; set; }
    }

    /// <summary>
    /// Estrutura de retorno da tentativa de aplicar uma nova elevacao.
    /// </summary>
    private readonly record struct ResultadoAjustePonto(bool Sucesso, string Mensagem)
    {
        public static ResultadoAjustePonto ComSucesso() => new(true, string.Empty);

        public static ResultadoAjustePonto Falha(string mensagem) => new(false, mensagem);
    }
}
