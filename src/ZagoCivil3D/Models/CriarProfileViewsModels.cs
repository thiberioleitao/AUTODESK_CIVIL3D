using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para criacao automatica de profile views
/// a partir dos alinhamentos do desenho. Replica a rotina Dynamo 1.3.
/// </summary>
public sealed class CriarProfileViewsRequest
{
    /// <summary>Coordenada X (Easting) do ponto de insercao do primeiro profile view.</summary>
    public double CoordenadaX { get; set; }

    /// <summary>Coordenada Y (Northing) do ponto de insercao do primeiro profile view.</summary>
    public double CoordenadaY { get; set; }

    /// <summary>Offset adicional aplicado entre profile views ao empilha-los verticalmente (em unidades do desenho).</summary>
    public double OffsetAdicional { get; set; } = 50.0;

    /// <summary>Nome do estilo do profile view. Vazio = usar o estilo padrao do desenho.</summary>
    public string NomeEstiloProfileView { get; set; } = string.Empty;

    /// <summary>Nome do conjunto de bands. Vazio = usar o conjunto padrao do desenho.</summary>
    public string NomeConjuntoBands { get; set; } = string.Empty;

    /// <summary>Sufixo adicionado ao nome do alinhamento para formar o nome do profile view.</summary>
    public string SufixoProfileView { get; set; } = "_PV";

    /// <summary>Quando verdadeiro, apaga profile views existentes com o mesmo nome antes de criar.</summary>
    public bool SubstituirExistentes { get; set; } = false;
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando e na janela.
/// </summary>
public sealed class CriarProfileViewsResultado
{
    /// <summary>Total de alinhamentos encontrados no desenho.</summary>
    public int TotalAlinhamentos { get; set; }

    /// <summary>Quantidade de profile views criados com sucesso.</summary>
    public int TotalProfileViewsCriados { get; set; }

    /// <summary>Quantidade de alinhamentos ignorados (ja possuem PV com o mesmo nome e nao substituir).</summary>
    public int TotalProfileViewsIgnorados { get; set; }

    /// <summary>Quantidade de profile views existentes que foram substituidos.</summary>
    public int TotalProfileViewsSubstituidos { get; set; }

    /// <summary>Nomes dos profile views criados.</summary>
    public List<string> NomesCriados { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
