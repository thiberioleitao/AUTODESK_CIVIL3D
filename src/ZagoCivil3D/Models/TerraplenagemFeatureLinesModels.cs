using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Reune todos os parametros preenchidos na UI para executar
/// o fluxo equivalente ao script Dynamo de terraplenagem.
/// </summary>
public sealed class TerraplenagemFeatureLinesRequest
{
    /// <summary>
    /// Nome do site de onde as feature lines serao lidas.
    /// </summary>
    public string NomeSite { get; set; } = string.Empty;

    /// <summary>
    /// Nome da layer que contem o poligono de filtro.
    /// </summary>
    public string NomeCamadaPoligono { get; set; } = string.Empty;

    /// <summary>
    /// Nome da superficie base usada na etapa de ajuste de terraplenagem.
    /// </summary>
    public string NomeSuperficieBase { get; set; } = string.Empty;

    /// <summary>
    /// Limite absoluto da deflexao aceitavel entre os segmentos adjacentes.
    /// </summary>
    public double DeflexaoLimite { get; set; } = 0.0129;

    /// <summary>
    /// Quantidade maxima de PI points analisados por feature line.
    /// </summary>
    public int NumeroMaximoPontosPorFeatureLine { get; set; } = 1000;

    /// <summary>
    /// Passo incremental aplicado no ajuste de elevacao dos pontos.
    /// </summary>
    public double PassoIncrementalAjusteCota { get; set; } = 0.01;

    /// <summary>
    /// Quantidade maxima de passadas globais no ajuste de deflexao.
    /// </summary>
    public int QuantidadePassadasGlobais { get; set; } = 2;

    /// <summary>
    /// Percentual minimo de pontos adequados para encerrar a passada global.
    /// </summary>
    public double PercentualObjetivo { get; set; } = 1.0;

    /// <summary>
    /// Quantidade maxima de tentativas locais por ponto ou segmento.
    /// </summary>
    public int NumeroTentativasPorPonto { get; set; } = 100;

    /// <summary>
    /// Diferenca minima entre FL e superficie para considerar um PI como ja alterado.
    /// </summary>
    public double ToleranciaBaixaSuperficie { get; set; } = 0.01;

    /// <summary>
    /// Diferenca maxima aceitavel entre FL e superficie no meio do segmento.
    /// </summary>
    public double ToleranciaAltaSuperficie { get; set; } = 0.10;

    /// <summary>
    /// Quantidade maxima de segmentos criticos ajustados por feature line.
    /// </summary>
    public int MaximoAjustesPorFeatureLine { get; set; } = 1;
}

/// <summary>
/// Consolida o resumo da execucao para exibicao na linha de comando.
/// </summary>
public sealed class TerraplenagemFeatureLinesResultado
{
    /// <summary>
    /// Quantidade total de feature lines recuperadas do site.
    /// </summary>
    public int TotalFeatureLinesNoSite { get; set; }

    /// <summary>
    /// Quantidade de feature lines que passaram no filtro do poligono.
    /// </summary>
    public int TotalFeatureLinesFiltradas { get; set; }

    /// <summary>
    /// Quantidade de feature lines que tiveram a etapa de deflexao executada.
    /// </summary>
    public int TotalFeatureLinesComDeflexaoAjustada { get; set; }

    /// <summary>
    /// Quantidade de feature lines que tiveram a etapa de superficie executada.
    /// </summary>
    public int TotalFeatureLinesComSuperficieAjustada { get; set; }

    /// <summary>
    /// Logs detalhados de processamento por feature line.
    /// </summary>
    public List<string> Logs { get; } = new();

    /// <summary>
    /// Mensagens de erro ou de validacao acumuladas na execucao.
    /// </summary>
    public List<string> MensagensErro { get; } = new();
}
