using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para criar catchments a partir das hatches
/// de uma layer. Replica a rotina Dynamo
/// "CRIAR CATCHMENT A PARTIR DAS HATCHS".
///
/// Fluxo resumido:
///   1. Le todas as hatches da layer informada e usa o contorno de cada uma
///      como boundary do catchment.
///   2. Para cada hatch, procura um MText (na layer de IDs) cujo ponto de
///      insercao esteja dentro do contorno — o conteudo do MText e o
///      identificador da subbacia.
///   3. Para cada hatch, procura uma polilinha (na layer de talvegues) cujos
///      extremos estejam ambos contidos no contorno — essa polilinha vira
///      o flow path do catchment.
///   4. Cria o catchment dentro do grupo informado, com nome
///      "{PrefixoBacia}-{idSubbacia}".
/// </summary>
public sealed class CriarCatchmentsDeHatchsRequest
{
    /// <summary>Nome da layer onde estao as hatches que definem as subbacias.</summary>
    public string NomeLayerHatches { get; set; } = string.Empty;

    /// <summary>Nome da layer dos MTexts com os IDs das subbacias.</summary>
    public string NomeLayerMTextsIds { get; set; } = string.Empty;

    /// <summary>Nome da layer das polilinhas de talvegue (flow path).</summary>
    public string NomeLayerTalvegues { get; set; } = string.Empty;

    /// <summary>Nome do Catchment Group onde os catchments serao adicionados. Se nao existir, sera criado.</summary>
    public string NomeGrupoCatchment { get; set; } = string.Empty;

    /// <summary>Prefixo da bacia, usado como prefixo do nome de cada catchment. Ex.: "SB".</summary>
    public string PrefixoBacia { get; set; } = "SB";

    /// <summary>Separador entre o prefixo e o ID da subbacia. Default "-".</summary>
    public string SeparadorNome { get; set; } = "-";

    /// <summary>Nome da superficie de referencia usada pelo catchment para calcular flow path.</summary>
    public string NomeSuperficie { get; set; } = string.Empty;

    /// <summary>Nome do estilo de catchment. Se vazio, usa o primeiro estilo disponivel.</summary>
    public string NomeEstiloCatchment { get; set; } = string.Empty;

    /// <summary>
    /// Se verdadeiro, a polilinha de talvegue contida na hatch sera usada
    /// como flow path do catchment. Caso contrario, apenas o contorno e
    /// criado (sem flow path).
    /// </summary>
    public bool ConfigurarFlowPath { get; set; } = true;

    /// <summary>
    /// Se verdadeiro, remove catchments com o mesmo nome antes de criar o novo.
    /// Evita conflito de unicidade do nome dentro do grupo.
    /// </summary>
    public bool SubstituirExistentes { get; set; } = true;
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando e na janela.
/// </summary>
public sealed class CriarCatchmentsDeHatchsResultado
{
    /// <summary>Quantidade de hatches encontradas na layer informada.</summary>
    public int TotalHatchesEncontradas { get; set; }

    /// <summary>Quantidade de MTexts encontrados na layer de IDs.</summary>
    public int TotalMTextsEncontrados { get; set; }

    /// <summary>Quantidade de polilinhas encontradas na layer de talvegues.</summary>
    public int TotalTalveguesEncontrados { get; set; }

    /// <summary>Quantidade de catchments criados com sucesso.</summary>
    public int TotalCatchmentsCriados { get; set; }

    /// <summary>Quantidade de catchments criados com flow path atribuido.</summary>
    public int TotalComFlowPath { get; set; }

    /// <summary>Quantidade de catchments substituidos (existentes com o mesmo nome).</summary>
    public int TotalSubstituidos { get; set; }

    /// <summary>Nomes dos catchments criados.</summary>
    public List<string> NomesCriados { get; } = new();

    /// <summary>Hatches ignoradas porque nao foi possivel resolver MText/talvegue.</summary>
    public List<string> AvisosHatchesIgnoradas { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
