using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para criar regioes em todos os corredores do
/// desenho, replicando a rotina Dynamo
/// "1.5 - CRIAR REGIOES A PARTIR DE CORREDORES".
///
/// As estacoes de quebra em cada baseline sao derivadas de tres criterios:
/// (a) cruzamentos com outros alinhamentos;
/// (b) mudancas bruscas de direcao (angulo entre segmentos do alinhamento);
/// (c) mudancas bruscas de declividade (variacao entre trechos do perfil).
/// </summary>
public sealed class CriarRegioesCorredoresRequest
{
    /// <summary>Nome do Assembly a usar nas regioes criadas.</summary>
    public string NomeAssembly { get; set; } = string.Empty;

    // ---------- Criterios de geracao de estacoes ----------

    /// <summary>
    /// Limite de mudanca de direcao (graus) entre segmentos consecutivos do
    /// alinhamento. Quando a deflexao supera esse valor, a estacao do vertice
    /// vira ponto de quebra de regiao. Padrao 60 graus.
    /// </summary>
    public double LimiteAnguloDirecao { get; set; } = 60.0;

    /// <summary>
    /// Variacao admissivel de declividade entre segmentos consecutivos do
    /// perfil de projeto. Variacoes maiores que esse valor geram estacoes de
    /// quebra. Padrao 0.05 (adimensional).
    /// </summary>
    public double LimiteMudancaDeclividade { get; set; } = 0.05;

    /// <summary>
    /// Espacamento minimo entre estacoes criticas de declividade para evitar
    /// regioes muito pequenas decorrentes de varios PVIs proximos. Padrao 50 m.
    /// </summary>
    public double EspacamentoMinimoDeclividade { get; set; } = 50.0;

    /// <summary>
    /// Espacamento minimo entre estacoes criticas (direcao/declividade) e os
    /// cruzamentos. Evita duplicar quebras muito proximas. Padrao 50 m.
    /// </summary>
    public double EspacamentoMinimoCruzamentos { get; set; } = 50.0;

    // ---------- Frequencia horizontal ----------

    /// <summary>Frequencia de insercao em trechos tangentes (m). Padrao 1.</summary>
    public double FrequenciaTangente { get; set; } = 1.0;

    /// <summary>Incremento ao longo de curvas horizontais (m). Padrao 0.5.</summary>
    public double IncrementoCurvaHorizontal { get; set; } = 0.5;

    /// <summary>Distancia mid-ordinate para tesselacao de curvas. Padrao 1.</summary>
    public double DistanciaMidOrdinate { get; set; } = 1.0;

    /// <summary>Frequencia em espirais (m). Padrao 1.</summary>
    public double FrequenciaEspiral { get; set; } = 1.0;

    /// <summary>Insere assemblies nos geometry points do alinhamento (horizontal).</summary>
    public bool HorizontalEmGeometryPoints { get; set; } = true;

    /// <summary>Insere em pontos criticos de superelevacao.</summary>
    public bool EmPontosSuperelevacao { get; set; } = true;

    // ---------- Frequencia vertical ----------

    /// <summary>Frequencia em curvas verticais (m). Padrao 1.</summary>
    public double FrequenciaCurvaVertical { get; set; } = 1.0;

    /// <summary>Insere em geometry points do perfil.</summary>
    public bool VerticalEmGeometryPoints { get; set; } = true;

    /// <summary>Insere em pontos de minimo/maximo do perfil.</summary>
    public bool VerticalEmHighLowPoints { get; set; } = true;

    // ---------- Frequencia offset target ----------

    /// <summary>Insere nos geometry points de offset targets.</summary>
    public bool OffsetEmGeometryPoints { get; set; } = true;

    /// <summary>
    /// Insere adjacente ao start/end do offset target (quando falso, insere
    /// exatamente no start/end).
    /// </summary>
    public bool OffsetAdjacenteStartEnd { get; set; } = true;

    /// <summary>Incremento ao longo de curvas de offset target (m). Padrao 1.</summary>
    public double IncrementoCurvaOffset { get; set; } = 1.0;

    /// <summary>Distancia mid-ordinate para offset target. Padrao 1.</summary>
    public double DistanciaMidOrdinateOffset { get; set; } = 1.0;
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando e na janela.
/// </summary>
public sealed class CriarRegioesCorredoresResultado
{
    public int TotalCorredores { get; set; }
    public int TotalBaselines { get; set; }
    public int TotalRegioesCriadas { get; set; }
    public int TotalBaselinesProcessadas { get; set; }

    /// <summary>Nomes (corridor/baseline -> regiao) para conferencia.</summary>
    public List<string> NomesRegioesCriadas { get; } = new();

    /// <summary>Logs detalhados (um por baseline).</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
