using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros preenchidos na UI para aplicar um Label Set Style a todos os
/// alinhamentos do desenho. Replica a rotina Dynamo
/// "MUDAR LABEL SET DOS ALINHAMENTOS".
/// </summary>
public sealed class MudarLabelSetAlinhamentosRequest
{
    /// <summary>Nome do Alignment Label Set Style a aplicar.</summary>
    public string NomeLabelSetStyle { get; set; } = string.Empty;

    /// <summary>
    /// Quando verdadeiro, apaga todos os label groups existentes dos
    /// alinhamentos antes de importar o novo set. Evita que labels antigos
    /// fiquem misturados com os novos.
    /// </summary>
    public bool ApagarExistentes { get; set; } = true;
}

/// <summary>
/// Resumo da execucao para exibicao na linha de comando e na janela.
/// </summary>
public sealed class MudarLabelSetAlinhamentosResultado
{
    /// <summary>Total de alinhamentos encontrados no desenho.</summary>
    public int TotalAlinhamentos { get; set; }

    /// <summary>Quantidade de alinhamentos cujo label set foi aplicado com sucesso.</summary>
    public int TotalAplicadosComSucesso { get; set; }

    /// <summary>Quantidade de label groups antigos apagados ao longo do processo.</summary>
    public int TotalLabelGroupsApagados { get; set; }

    /// <summary>Nomes dos alinhamentos processados com sucesso.</summary>
    public List<string> NomesProcessados { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
