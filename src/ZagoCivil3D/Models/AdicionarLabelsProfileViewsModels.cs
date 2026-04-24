using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para adicionar Profile Line Labels em todos
/// os profile views do desenho. Replica a rotina Dynamo
/// "ADICIONAR LABELS AS PROFILE VIEWS": para cada alinhamento, itera por seus
/// profile views e, dentre os profiles associados, seleciona aqueles cujo
/// nome contem o texto do filtro; em cada combinacao valida cria um
/// ProfileLineLabelGroup com o Profile Line Label Style escolhido.
/// </summary>
public sealed class AdicionarLabelsProfileViewsRequest
{
    /// <summary>Nome do Profile Line Label Style a aplicar (ex.: "Percent Grade").</summary>
    public string NomeProfileLineLabelStyle { get; set; } = string.Empty;

    /// <summary>
    /// Texto usado para filtrar os profiles pelo nome (case-insensitive, contains).
    /// Apenas os profiles cujo nome contem este texto recebem labels. Ex.: "TER".
    /// </summary>
    public string FiltroNomeProfile { get; set; } = string.Empty;
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando e na janela.
/// </summary>
public sealed class AdicionarLabelsProfileViewsResultado
{
    /// <summary>Total de alinhamentos encontrados no desenho.</summary>
    public int TotalAlinhamentos { get; set; }

    /// <summary>Total de profile views encontrados em todos os alinhamentos.</summary>
    public int TotalProfileViews { get; set; }

    /// <summary>Total de profiles que passaram no filtro de nome.</summary>
    public int TotalProfilesFiltrados { get; set; }

    /// <summary>Quantidade de label groups criados com sucesso.</summary>
    public int TotalLabelsCriados { get; set; }

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
