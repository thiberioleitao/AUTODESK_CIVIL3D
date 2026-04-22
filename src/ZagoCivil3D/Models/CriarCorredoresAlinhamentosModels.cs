using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para criar corredores vazios a partir de todos
/// os alinhamentos do desenho. Replica a rotina Dynamo
/// "CRIAR CORREDORES A PARTIR DE ALINHAMENTOS": um corredor por alinhamento,
/// nomeado com base no nome do alinhamento.
/// </summary>
public sealed class CriarCorredoresAlinhamentosRequest
{
    /// <summary>
    /// Prefixo opcional aplicado ao nome do corredor. Ex.: prefixo "CORR_"
    /// + alinhamento "EIXO-01" gera corredor "CORR_EIXO-01". Vazio por padrao
    /// para reproduzir o Dynamo (usa o nome do alinhamento tal como esta).
    /// </summary>
    public string PrefixoNome { get; set; } = string.Empty;

    /// <summary>
    /// Sufixo opcional aplicado ao nome do corredor. Vazio por padrao.
    /// </summary>
    public string SufixoNome { get; set; } = string.Empty;

    /// <summary>
    /// Quando verdadeiro, alinhamentos cujo nome final ja possui um corredor
    /// com o mesmo nome sao ignorados silenciosamente. Quando falso, o
    /// servico tenta criar e reporta erro em caso de colisao.
    /// </summary>
    public bool IgnorarExistentes { get; set; } = true;
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando e na janela.
/// </summary>
public sealed class CriarCorredoresAlinhamentosResultado
{
    /// <summary>Total de alinhamentos encontrados no desenho.</summary>
    public int TotalAlinhamentos { get; set; }

    /// <summary>Quantidade de corredores criados com sucesso.</summary>
    public int TotalCorredoresCriados { get; set; }

    /// <summary>Quantidade de alinhamentos ignorados por ja possuirem corredor.</summary>
    public int TotalIgnorados { get; set; }

    /// <summary>Nomes dos corredores criados.</summary>
    public List<string> NomesCriados { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
