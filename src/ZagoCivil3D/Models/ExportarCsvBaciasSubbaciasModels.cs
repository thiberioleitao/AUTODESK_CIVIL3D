using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Parametros da exportacao CSV de bacias/subbacias. Replica a rotina Dynamo
/// "EXPORTAR CSV BACIAS E SUBBACIAS-ID, AREA, TALVEGUE, DECLIVIDADE E ID_JUSANTE".
///
/// Fluxo resumido:
///   1. Varre todas as hatches cujo layer comece com o <see cref="PrefixoLayersHatches"/>.
///      Cada hatch representa uma subbacia e o layer identifica a bacia "pai".
///   2. Extrai o contorno da hatch; procura o MText da layer
///      <see cref="NomeLayerMTextsSubbacias"/> contido no contorno (conteudo com
///      <see cref="PrefixoTextoSubbacia"/>) — isso define o ID da SUBBACIA.
///   3. Procura a polyline da layer <see cref="NomeLayerTalvegues"/> cujos dois
///      extremos estejam dentro do contorno — isso da o TALVEGUE.
///   4. Calcula L_TALVEGUE (comprimento) e S_TALVEGUE (declividade media
///      (zm - zj) / L a partir de <see cref="NomeSuperficie"/>).
///   5. Projeta o ponto final do talvegue em cada regiao dos corredores; a
///      primeira cujo eixo passa dentro de <see cref="RaioBuscaJusante"/>
///      metros do ponto final informa o ID_JUSANTE.
///   6. Monta o CSV no caminho <see cref="CaminhoCsv"/> com as colunas
///      BACIA, SUBBACIA, AREA, L_TALVEGUE, S_TALVEGUE, ID_JUSANTE e 4 verificacoes.
/// </summary>
public sealed class ExportarCsvBaciasSubbaciasRequest
{
    /// <summary>Caminho absoluto do arquivo CSV a gerar (sera sobrescrito).</summary>
    public string CaminhoCsv { get; set; } = string.Empty;

    /// <summary>
    /// Layer onde estao os MTexts dos IDs das subbacias. Padrao Zago: "A - TEXT BACIA".
    /// </summary>
    public string NomeLayerMTextsSubbaciasId { get; set; } = "A - TEXT BACIA";

    /// <summary>
    /// Substring que o conteudo do MText precisa conter para ser considerado
    /// um identificador de subbacia. Padrao Zago: "SB".
    /// </summary>
    public string PrefixoTextoSubbacia { get; set; } = "SB";

    /// <summary>
    /// Layer onde estao as polilinhas de talvegue (uma por subbacia). Padrao Zago:
    /// "HDR - TALVEGUES SUBBACIAS".
    /// </summary>
    public string NomeLayerTalvegues { get; set; } = "HDR - TALVEGUES SUBBACIAS";

    /// <summary>
    /// Prefixo/substring que deve aparecer no nome das layers de hatches das
    /// subbacias. Todas as layers cujo nome contem este texto sao varridas.
    /// Padrao Zago: "HDR-BACIA" — ex.: casa "HDR-BACIA 01", "HDR-BACIA 02".
    /// </summary>
    public string PrefixoLayersHatches { get; set; } = "HDR-BACIA";

    /// <summary>
    /// Prefixo removido do nome da layer para compor o valor da coluna BACIA.
    /// O que sobra e aplicado em title case; ex.: layer "HDR-BACIA 01", com
    /// prefixo "HDR-", gera BACIA "Bacia 01". Deixe vazio para usar o nome
    /// completo da layer.
    /// </summary>
    public string PrefixoRemoverDoNomeBacia { get; set; } = "HDR-";

    /// <summary>Nome da superficie TIN usada para calcular cotas do talvegue.</summary>
    public string NomeSuperficie { get; set; } = string.Empty;

    /// <summary>
    /// Raio de busca (em metros do desenho) a partir do ponto final do talvegue
    /// para encontrar a regiao do corredor correspondente ao ID_JUSANTE. Padrao: 10.
    /// </summary>
    public double RaioBuscaJusante { get; set; } = 10.0;
}

/// <summary>
/// Uma linha do CSV exportado. E usada tanto para gerar o arquivo quanto para
/// alimentar a DataGrid da aba "Resultados" da janela.
/// </summary>
public sealed class LinhaCsvBaciasSubbacias
{
    /// <summary>Texto da coluna "OK" quando nao ha aviso de verificacao.</summary>
    public const string VerfOk = "OK";

    public string Bacia { get; set; } = string.Empty;
    public string Subbacia { get; set; } = string.Empty;
    public double? Area { get; set; }
    public double? Comprimento { get; set; }
    public double? Declividade { get; set; }
    public string IdJusante { get; set; } = string.Empty;
    public string VerfSb { get; set; } = VerfOk;
    public string VerfA { get; set; } = VerfOk;
    public string VerfL { get; set; } = VerfOk;
    public string VerfS { get; set; } = VerfOk;

    /// <summary>True se qualquer verificacao VERF_* for diferente de "OK".</summary>
    public bool TemAviso =>
        VerfSb != VerfOk || VerfA != VerfOk || VerfL != VerfOk || VerfS != VerfOk;
}

/// <summary>
/// Resumo da execucao, exibido na linha de comando e na janela.
/// </summary>
public sealed class ExportarCsvBaciasSubbaciasResultado
{
    /// <summary>Total de hatches encontradas (uma por subbacia).</summary>
    public int TotalHatches { get; set; }

    /// <summary>Total de MTexts de IDs encontrados na layer configurada.</summary>
    public int TotalMTextsIds { get; set; }

    /// <summary>Total de polilinhas de talvegue encontradas.</summary>
    public int TotalTalvegues { get; set; }

    /// <summary>Total de baseline regions encontradas nos corredores.</summary>
    public int TotalRegioesCorredores { get; set; }

    /// <summary>Linhas efetivamente escritas no CSV (apos o header).</summary>
    public int TotalLinhasCsv { get; set; }

    /// <summary>Quantidade de subbacias com pelo menos um VERF diferente de "OK".</summary>
    public int TotalLinhasComAviso { get; set; }

    /// <summary>Caminho completo do CSV gerado.</summary>
    public string CaminhoCsvGerado { get; set; } = string.Empty;

    /// <summary>
    /// Linhas exportadas (com os mesmos dados do CSV). Usadas pela UI para
    /// popular a aba "Resultados" com DataGrid e lista de avisos.
    /// </summary>
    public List<LinhaCsvBaciasSubbacias> Linhas { get; } = new();

    /// <summary>Logs detalhados de processamento.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro acumuladas.</summary>
    public List<string> MensagensErro { get; } = new();
}
