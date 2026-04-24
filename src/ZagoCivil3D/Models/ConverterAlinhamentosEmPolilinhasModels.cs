using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para converter alinhamentos do Civil 3D em
/// polilinhas 2D do AutoCAD (entidade <c>Polyline</c>).
/// </summary>
public sealed class ConverterAlinhamentosEmPolilinhasRequest
{
    /// <summary>
    /// Nome da layer aplicada as polilinhas criadas. Caso vazio, as polilinhas
    /// usam a layer corrente do desenho. Se a layer nao existir, ela e criada.
    /// </summary>
    public string NomeLayer { get; set; } = "C-ROAD-POLY";

    /// <summary>
    /// Filtro de nome (substring case-insensitive) aplicado aos alinhamentos
    /// do desenho. Apenas alinhamentos cujo nome contem este texto sao
    /// convertidos. Caso vazio, todos os alinhamentos sao processados.
    /// Ex.: "Acesso" seleciona "Acesso 01", "Acesso Principal", etc.
    /// </summary>
    public string FiltroNome { get; set; } = string.Empty;

    /// <summary>
    /// Passo de discretizacao usado para aproximar espirais (clotoides) em
    /// segmentos retos. Valores menores geram polilinhas mais fieis as curvas
    /// de transicao, ao custo de mais vertices. Expresso em metros.
    /// </summary>
    public double PassoDiscretizacaoEspirais { get; set; } = 1.0;

    /// <summary>
    /// Quando verdadeiro, arcos do alinhamento sao preservados na polilinha
    /// por meio do <c>bulge</c> do vertice. Quando falso, os arcos tambem
    /// sao discretizados em segmentos retos pelo mesmo passo das espirais.
    /// </summary>
    public bool PreservarArcos { get; set; } = true;

    /// <summary>
    /// Elevation (Z constante) aplicada as polilinhas 2D criadas. A entidade
    /// <see cref="Autodesk.AutoCAD.DatabaseServices.Polyline"/> e plana; este
    /// valor define a altura do plano que a contem.
    /// </summary>
    public double Elevation { get; set; } = 0.0;

    /// <summary>
    /// Quando verdadeiro, os alinhamentos originais sao apagados apos a
    /// criacao das polilinhas. Acao destrutiva: padrao desligado.
    /// </summary>
    public bool ApagarAlinhamentosOriginais { get; set; } = false;

    /// <summary>
    /// Quando verdadeiro, apenas calcula e reporta a conversao sem alterar
    /// o desenho. Util para validar os parametros antes da execucao.
    /// </summary>
    public bool DryRun { get; set; } = false;
}

/// <summary>
/// Informacoes sobre a conversao de um alinhamento individual.
/// </summary>
public sealed class AlinhamentoConvertido
{
    public string NomeAlinhamento { get; set; } = string.Empty;
    public int TotalVertices { get; set; }
    public int TotalLinhas { get; set; }
    public int TotalArcos { get; set; }
    public int TotalEspirais { get; set; }
    public int VerticesDiscretizadosEspirais { get; set; }
    public double ComprimentoAlinhamento { get; set; }
}

/// <summary>
/// Resumo da execucao da conversao para exibicao na linha de comando e
/// na janela.
/// </summary>
public sealed class ConverterAlinhamentosEmPolilinhasResultado
{
    /// <summary>Total de alinhamentos encontrados no desenho.</summary>
    public int TotalAlinhamentos { get; set; }

    /// <summary>Alinhamentos que passaram pelo filtro de nome.</summary>
    public int TotalAlinhamentosFiltrados { get; set; }

    /// <summary>Alinhamentos convertidos com sucesso em polilinha.</summary>
    public int TotalPolilinhasCriadas { get; set; }

    /// <summary>Alinhamentos originais apagados apos a conversao.</summary>
    public int TotalAlinhamentosApagados { get; set; }

    /// <summary>Detalhamento por alinhamento convertido.</summary>
    public List<AlinhamentoConvertido> Convertidos { get; } = new();

    /// <summary>Avisos nao fatais durante a conversao.</summary>
    public List<string> Avisos { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
