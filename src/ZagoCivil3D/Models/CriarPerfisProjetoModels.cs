using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para criacao automatica de perfis de projeto.
/// </summary>
public sealed class CriarPerfisProjetoRequest
{
    /// <summary>Nome da superficie TIN usada para amostragem de elevacoes.</summary>
    public string NomeSuperficie { get; set; } = string.Empty;

    /// <summary>Nome do estilo de perfil a ser aplicado.</summary>
    public string NomeEstiloPerfil { get; set; } = string.Empty;

    /// <summary>Nome do conjunto de rotulos do perfil.</summary>
    public string NomeConjuntoRotulosPerfil { get; set; } = string.Empty;

    /// <summary>Nome da layer de destino para os perfis criados.</summary>
    public string NomeCamada { get; set; } = string.Empty;

    /// <summary>Intervalo entre estacoes amostradas, em metros.</summary>
    public double EspacamentoEstacas { get; set; } = 10.0;

    /// <summary>Offset vertical subtraido da cota do terreno.</summary>
    public double OffsetAbaixoTerreno { get; set; } = 0.100;

    /// <summary>Declividade minima imposta ao perfil de projeto (m/m).</summary>
    public double DeclivedadeMinima { get; set; } = 0.002;

    /// <summary>Variacao admissivel de declividade para suavizacao de PIs.</summary>
    public double LimiteVariacaoDeclividade { get; set; } = 0.050;

    /// <summary>Sufixo adicionado ao nome do alinhamento para formar o nome do perfil.</summary>
    public string SufixoPerfil { get; set; } = "_PROJETO";

    /// <summary>Quando verdadeiro, apaga perfis existentes com o mesmo nome antes de criar.</summary>
    public bool SubstituirPerfisExistentes { get; set; } = false;
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando.
/// </summary>
public sealed class CriarPerfisProjetoResultado
{
    /// <summary>Total de alinhamentos encontrados no desenho.</summary>
    public int TotalAlinhamentos { get; set; }

    /// <summary>Quantidade de perfis criados com sucesso.</summary>
    public int TotalPerfisCriados { get; set; }

    /// <summary>Quantidade de alinhamentos ignorados (perfil ja existente, sem pontos, etc.).</summary>
    public int TotalPerfisIgnorados { get; set; }

    /// <summary>Quantidade de perfis existentes que foram substituidos.</summary>
    public int TotalPerfisSubstituidos { get; set; }

    /// <summary>Nomes dos perfis criados.</summary>
    public List<string> NomesCriados { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
