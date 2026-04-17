using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para criacao automatica de surface profiles
/// (perfis do terreno natural e de terraplenagem) a partir de uma superficie TIN.
/// Replica a rotina Dynamo 1.1 "CRIAR PERFIS DO TERRENO NATURAL E TERRAPLENAGEM
/// A PARTIR DE ALINHAMENTOS".
/// </summary>
public sealed class CriarPerfisTerrenoRequest
{
    /// <summary>Nome da superficie TIN usada para amostrar o terreno.</summary>
    public string NomeSuperficie { get; set; } = string.Empty;

    /// <summary>Nome do estilo de perfil a ser aplicado aos perfis criados.</summary>
    public string NomeEstiloPerfil { get; set; } = string.Empty;

    /// <summary>Nome do conjunto de rotulos do perfil.</summary>
    public string NomeConjuntoRotulosPerfil { get; set; } = string.Empty;

    /// <summary>Nome da layer de destino para os perfis criados.</summary>
    public string NomeCamada { get; set; } = string.Empty;

    /// <summary>Quando verdadeiro, apaga perfis existentes com o mesmo nome antes de criar.</summary>
    public bool SubstituirPerfisExistentes { get; set; } = false;
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando e na barra de status da janela.
/// </summary>
public sealed class CriarPerfisTerrenoResultado
{
    /// <summary>Total de alinhamentos encontrados no desenho.</summary>
    public int TotalAlinhamentos { get; set; }

    /// <summary>Quantidade de perfis criados com sucesso.</summary>
    public int TotalPerfisCriados { get; set; }

    /// <summary>Quantidade de alinhamentos ignorados (perfil ja existente e nao substituir).</summary>
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
