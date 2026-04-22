using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para criar o talvegue (FlowPath) dos
/// catchments existentes a partir das polilinhas desenhadas em uma layer
/// especifica. Replica a rotina Dynamo
/// "CRIAR TALVEGUE DOS CATCHMENTS A PARTIR DAS POLILINHAS".
/// </summary>
public sealed class CriarTalveguesCatchmentsRequest
{
    /// <summary>Layer onde estao as polilinhas que representam os talvegues.</summary>
    public string NomeCamadaPolilinhas { get; set; } = string.Empty;

    /// <summary>
    /// Quando verdadeiro, sobrescreve o FlowPath do catchment mesmo que ja
    /// exista algum segmento definido. Quando falso, catchments que ja possuem
    /// FlowPath sao preservados e apenas logados.
    /// </summary>
    public bool SubstituirFlowPathExistente { get; set; } = true;

    /// <summary>
    /// Tolerancia XY (em unidades do desenho) para aceitar um ponto sobre a
    /// borda do boundary de um catchment como "contido". Evita que talvegues
    /// que se encostam no limite sejam rejeitados por diferencas numericas.
    /// </summary>
    public double ToleranciaBorda { get; set; } = 1e-6;
}

/// <summary>
/// Resumo da execucao para exibir na linha de comando do AutoCAD e na
/// barra de status da janela.
/// </summary>
public sealed class CriarTalveguesCatchmentsResultado
{
    /// <summary>Polilinhas encontradas na layer configurada.</summary>
    public int TotalPolilinhas { get; set; }

    /// <summary>Catchments existentes no desenho.</summary>
    public int TotalCatchments { get; set; }

    /// <summary>Catchments que tiveram o FlowPath criado agora.</summary>
    public int TotalFlowPathsCriados { get; set; }

    /// <summary>Catchments cujo FlowPath foi substituido.</summary>
    public int TotalFlowPathsSubstituidos { get; set; }

    /// <summary>Catchments ignorados porque ja possuiam FlowPath e o usuario nao escolheu substituir.</summary>
    public int TotalCatchmentsIgnorados { get; set; }

    /// <summary>Polilinhas que nao puderam ser associadas a nenhum catchment (a revisar).</summary>
    public int TotalPolilinhasSemCatchment { get; set; }

    /// <summary>Nomes dos catchments atualizados com sucesso.</summary>
    public List<string> NomesCatchmentsAtualizados { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas durante a execucao.</summary>
    public List<string> MensagensErro { get; } = new();
}
