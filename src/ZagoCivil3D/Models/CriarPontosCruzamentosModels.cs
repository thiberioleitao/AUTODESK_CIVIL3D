using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para criar CogoPoints nos cruzamentos entre
/// alinhamentos. Porta da rotina Dynamo "CRIAR PONTOS NOS CRUZAMENTOS ENTRE
/// ALINHAMENTOS".
/// </summary>
public sealed class CriarPontosCruzamentosRequest
{
    /// <summary>
    /// Comprimento L do tracker, em metros. Define a altura (em Y) da janela
    /// que delimita cada bloco: um bloco contem pontos com Y entre o topo
    /// do bloco e <c>topo - ComprimentoTracker</c>.
    /// </summary>
    public double ComprimentoTracker { get; set; } = 140.0;

    /// <summary>
    /// Quantidade esperada de pontos por coluna dentro de cada bloco.
    /// Usado apenas para emitir avisos quando uma coluna nao tem o tamanho
    /// previsto. Nao altera a criacao dos pontos.
    /// </summary>
    public int PontosPorTracker { get; set; } = 15;

    /// <summary>
    /// Prefixo do rotulo sequencial aplicado como RawDescription dos
    /// CogoPoints. Ex.: prefixo "P" gera "P01", "P02", "P03"...
    /// </summary>
    public string Prefixo { get; set; } = "P";

    /// <summary>
    /// Quando verdadeiro, blocos impares comecam do lado oeste (mais ao
    /// noroeste) e blocos pares do lado leste (mais ao nordeste), de modo
    /// que a numeracao serpenteia entre blocos consecutivos. Quando falso,
    /// todos os blocos comecam do noroeste.
    /// </summary>
    public bool AlternarLadoInicial { get; set; } = false;

    /// <summary>
    /// Nome da layer aplicada aos CogoPoints criados. Caso vazio, os pontos
    /// usam a layer corrente do desenho.
    /// </summary>
    public string NomeLayer { get; set; } = "C-TOPO-TEXT";

    /// <summary>
    /// Tolerancia usada para considerar dois pontos como o mesmo cruzamento
    /// no plano XY. Distancias menores ou iguais a este valor sao colapsadas.
    /// </summary>
    public double ToleranciaXyDuplicados { get; set; } = 0.01;

    /// <summary>
    /// Tolerancia usada no agrupamento em colunas: pontos com <c>|X - Xref|</c>
    /// menor ou igual a este valor sao considerados parte da mesma coluna.
    /// </summary>
    public double ToleranciaXColuna { get; set; } = 0.10;

    /// <summary>
    /// Texto usado como descricao inicial dos CogoPoints antes da aplicacao
    /// do rotulo sequencial. Serve de fallback para pontos que, por algum
    /// motivo, nao recebam o rotulo final.
    /// </summary>
    public string Descricao { get; set; } = "CRUZAMENTO";

    /// <summary>
    /// Quando verdadeiro, os CogoPoints sao efetivamente criados no desenho.
    /// Quando falso, a rotina apenas calcula e reporta a ordenacao dos pontos
    /// sem alterar o desenho (util para validacoes em seco).
    /// </summary>
    public bool CriarCogoPoints { get; set; } = true;
}

/// <summary>
/// Informacoes de um bloco gerado durante o agrupamento dos cruzamentos.
/// </summary>
public sealed class BlocoCruzamento
{
    public int IndiceBloco { get; set; }

    /// <summary>Lado por onde o bloco comecou: "oeste" ou "leste".</summary>
    public string LadoInicial { get; set; } = "oeste";

    public double PontoBaseX { get; set; }
    public double PontoBaseY { get; set; }

    /// <summary>Quantidade de pontos ordenados dentro deste bloco.</summary>
    public int QuantidadePontos { get; set; }

    /// <summary>Rotulos sequenciais atribuidos aos pontos deste bloco.</summary>
    public List<string> Rotulos { get; } = new();
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando e na janela.
/// </summary>
public sealed class CriarPontosCruzamentosResultado
{
    /// <summary>Total de alinhamentos encontrados no desenho.</summary>
    public int TotalAlinhamentos { get; set; }

    /// <summary>Pares de alinhamentos testados para interseccao.</summary>
    public int TotalParesTestados { get; set; }

    /// <summary>Pares de alinhamentos com pelo menos um ponto de interseccao util.</summary>
    public int TotalParesComIntersecao { get; set; }

    /// <summary>Pontos brutos (antes da remocao de duplicados XY).</summary>
    public int TotalPontosBrutos { get; set; }

    /// <summary>Pontos unicos apos remocao de duplicados XY.</summary>
    public int TotalPontosUnicos { get; set; }

    /// <summary>Pontos efetivamente ordenados em blocos.</summary>
    public int TotalPontosOrdenados { get; set; }

    /// <summary>Quantidade de CogoPoints criados com sucesso.</summary>
    public int TotalCogoPointsCriados { get; set; }

    /// <summary>Quantidade de CogoPoints que tiveram a RawDescription aplicada.</summary>
    public int TotalRawDescriptionAplicadas { get; set; }

    /// <summary>Blocos gerados durante o agrupamento.</summary>
    public List<BlocoCruzamento> Blocos { get; } = new();

    /// <summary>Avisos nao fatais (ex.: coluna com tamanho inesperado).</summary>
    public List<string> Avisos { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
