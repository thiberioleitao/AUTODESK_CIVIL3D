using System.Collections.Generic;

namespace ZagoCivil3D.Models;

/// <summary>
/// Estrategia de escolha do PI point a ser ajustado quando uma violacao de deflexao e detectada.
/// </summary>
public enum ModoOtimizacaoDeflexao
{
    /// <summary>
    /// Ajusta sempre o PI central (indice i) do trio (i-1, i, i+1). Estrategia mais conservadora:
    /// a perturbacao fica restrita ao proprio ponto violador, sem propagar elevacao a vizinhos.
    /// </summary>
    PontoCentral,

    /// <summary>
    /// Avalia os tres candidatos do trio (Zi-1, Zi, Zi+1) e seleciona aquele com menor impacto
    /// volumetrico de terraplenagem (|dZ| ponderado pelo comprimento de influencia), com
    /// preferencia por estacoes ja alteradas na mesma feature line. Tende a produzir menos
    /// volume de corte/aterro e a concentrar as alteracoes em um numero menor de estacoes.
    /// </summary>
    MultiPonto
}

/// <summary>
/// Parametros preenchidos na UI para o ajuste iterativo de deflexao de feature lines.
/// </summary>
public sealed class AjustarDeflexaoFeatureLinesRequest
{
    /// <summary>
    /// Nome do site de onde as feature lines serao lidas. Quando vazio,
    /// todas as feature lines de todos os sites do desenho sao consideradas.
    /// </summary>
    public string NomeSite { get; set; } = string.Empty;

    /// <summary>
    /// Filtro opcional aplicado ao nome da feature line (substring case-insensitive).
    /// Quando vazio, todas as feature lines do site selecionado sao processadas.
    /// </summary>
    public string FiltroNome { get; set; } = string.Empty;

    /// <summary>
    /// Limite absoluto da deflexao aceitavel entre os segmentos adjacentes.
    /// Valor default 0.0129 equivale ao criterio usado na rotina Dynamo original.
    /// </summary>
    public double DeflexaoLimite { get; set; } = 0.0129;

    /// <summary>
    /// Passo minimo (piso) do ajuste iterativo de elevacao. Atua como a precisao minima do
    /// ajuste: o iterador parte deste passo e pode crescer ate <see cref="PassoMaximoDinamico"/>
    /// enquanto faz progresso, retornando a valores proximos deste piso ao bisectar no overshoot.
    /// </summary>
    public double PassoIncrementalAjusteCota { get; set; } = 0.01;

    /// <summary>
    /// Teto do passo dinamico. Limita o salto por iteracao para evitar overshoots grandes.
    /// Quando o ajuste faz progresso, o passo cresce ate este valor; ao detectar overshoot,
    /// o passo e reduzido pela metade (bisseccao) ate respeitar o limite de deflexao.
    /// </summary>
    public double PassoMaximoDinamico { get; set; } = 0.50;

    /// <summary>
    /// Indice maximo do PI avaliado <b>dentro de cada feature line</b>. Cada FL e processada
    /// independentemente, e em cada uma a varredura vai de PI 1 ate min(IndiceMaximo, n-2),
    /// onde n e o total de PIs dessa FL. Use valores altos (ex.: 1000) para nao limitar
    /// em FLs tipicas; use valores baixos para restringir o ajuste a um trecho inicial.
    /// </summary>
    public int IndiceMaximo { get; set; } = 1000;

    /// <summary>
    /// Numero de passadas globais sobre cada feature line. Passadas impares seguem no sentido
    /// montante -> jusante; passadas pares, no sentido inverso. Duas passadas sao o suficiente
    /// na maioria dos casos para acomodar efeitos de cascata do modo MultiPonto.
    /// </summary>
    public int QuantidadePassadas { get; set; } = 2;

    /// <summary>
    /// Quantidade maxima de iteracoes locais por ponto. Protege contra loops infinitos.
    /// </summary>
    public int NumeroMaximoIteracoesPorPonto { get; set; } = 10000;

    /// <summary>
    /// Estrategia de escolha do PI a ajustar. Ver <see cref="ModoOtimizacaoDeflexao"/>.
    /// </summary>
    public ModoOtimizacaoDeflexao ModoOtimizacao { get; set; } = ModoOtimizacaoDeflexao.MultiPonto;

    /// <summary>
    /// Fator multiplicador aplicado ao custo do candidato cuja estacao ja foi alterada
    /// em outra violacao da mesma feature line. Valor &lt; 1 favorece a reutilizacao
    /// (menos estacoes distintas modificadas). Usado apenas no modo MultiPonto.
    /// </summary>
    public double FatorReutilizacaoEstacao { get; set; } = 0.3;

    /// <summary>
    /// Quando verdadeiro, cria marcadores visuais (multileaders com seta apontando para o
    /// centro do PI) no desenho para as estacoes afetadas: cor azul para PIs ajustados com
    /// sucesso, cor vermelha para PIs em violacao remanescente ou falha geometrica.
    /// PIs que ja estavam dentro do limite nao recebem marcador.
    /// </summary>
    public bool CriarMarcadoresNoDesenho { get; set; } = true;

    /// <summary>
    /// Quando verdadeiro, apaga os marcadores criados por execucoes anteriores deste comando
    /// antes de criar os novos. A identificacao e feita via XData, de modo que nao interfere
    /// em outras entidades do desenho, mesmo na mesma layer.
    /// </summary>
    public bool LimparMarcadoresAnteriores { get; set; } = true;

    /// <summary>
    /// Layer dos marcadores (circulo verde + multileader verde) para PIs ajustados com
    /// sucesso. Quando vazia, usa <c>ZAGO_DEFLEXAO_AJUSTADO</c> (criada se nao existir).
    /// </summary>
    public string NomeLayerMarcadorAjustado { get; set; } = string.Empty;

    /// <summary>
    /// Layer dos marcadores (circulo vermelho + multileader vermelho) para PIs em violacao
    /// remanescente ou falha geometrica. Quando vazia, usa <c>ZAGO_DEFLEXAO_FALHA</c>.
    /// </summary>
    public string NomeLayerMarcadorFalha { get; set; } = string.Empty;

    /// <summary>
    /// Layer dos marcadores (apenas circulo azul, sem label) para PIs que ja estavam dentro
    /// do limite e nao precisaram de ajuste. Quando vazia, usa <c>ZAGO_DEFLEXAO_INALTERADO</c>.
    /// </summary>
    public string NomeLayerMarcadorInalterado { get; set; } = string.Empty;

    /// <summary>
    /// Altura do texto do multileader, em unidades do desenho. Tambem escala a seta
    /// proporcionalmente. Ajuste conforme a escala do desenho: valores maiores em mapas
    /// de larga escala, menores em desenhos detalhados.
    /// </summary>
    public double AlturaTextoMarcador { get; set; } = 1.0;
}

/// <summary>
/// Status final de um PI point avaliado, usado na tabela de relatorio.
/// </summary>
public enum StatusAjustePI
{
    /// <summary>PI ja estava dentro do limite antes da execucao; nenhuma alteracao feita.</summary>
    JaDentroDoLimite,

    /// <summary>PI violava o limite e ficou dentro do limite apos o ajuste.</summary>
    AjustadoComSucesso,

    /// <summary>PI continua violando o limite apos todas as passadas (nao convergiu).</summary>
    ViolacaoRemanescente,

    /// <summary>Nao foi possivel avaliar o PI (ex.: segmento adjacente com comprimento zero).</summary>
    FalhaGeometrica
}

/// <summary>
/// Status final de uma feature line inteira, usado na tabela de resumo.
/// </summary>
public enum StatusRelatorioFL
{
    /// <summary>Nenhuma violacao detectada (nem inicial, nem final).</summary>
    SemViolacoes,

    /// <summary>Todas as violacoes iniciais foram corrigidas.</summary>
    AjustadoComSucesso,

    /// <summary>Restaram violacoes apenas em PI de fronteira interna (PI 1 ou PI n-2).</summary>
    ViolacoesApenasEmFronteira,

    /// <summary>Restaram violacoes em PI internos (nao fronteira): requer intervencao manual.</summary>
    ViolacoesInternasRemanescentes,

    /// <summary>Houve falhas geometricas durante o processamento.</summary>
    Erro
}

/// <summary>
/// Relatorio de um unico PI point dentro de uma feature line, para exibicao na tabela.
/// </summary>
public sealed class RelatorioPontoPI
{
    /// <summary>Nome da feature line a que este PI pertence.</summary>
    public string NomeFeatureLine { get; set; } = string.Empty;

    /// <summary>Indice do PI dentro do conjunto PIPoint da feature line.</summary>
    public int IndicePI { get; set; }

    /// <summary>Quantidade total de PIs na feature line (para facilitar identificacao de fronteiras).</summary>
    public int TotalPIs { get; set; }

    /// <summary>Deflexao calculada antes de qualquer ajuste. Null quando geometricamente inavaliavel.</summary>
    public double? DeflexaoInicial { get; set; }

    /// <summary>Deflexao calculada apos todas as passadas. Null quando inavaliavel.</summary>
    public double? DeflexaoFinal { get; set; }

    /// <summary>
    /// Soma dos |dZ| aplicados a este PI quando ele foi alvo de ajuste (em qualquer passada).
    /// Zero indica que a estacao nao foi alterada.
    /// </summary>
    public double DeltaZAcumulado { get; set; }

    /// <summary>Quantas vezes esta estacao foi alvo de ajuste ao longo das passadas.</summary>
    public int VezesAlvo { get; set; }

    /// <summary>
    /// Soma das iteracoes internas do iterador quando esta estacao foi alvo.
    /// Reflete o "custo" de convergencia desta estacao (quanto menor, mais rapido convergiu).
    /// </summary>
    public int IteracoesAcumuladas { get; set; }

    /// <summary>Coordenada X (Easting) deste PI no desenho, apos o ajuste. Null quando geometricamente inavaliavel.</summary>
    public double? PosicaoX { get; set; }

    /// <summary>Coordenada Y (Northing) deste PI no desenho, apos o ajuste. Null quando geometricamente inavaliavel.</summary>
    public double? PosicaoY { get; set; }

    /// <summary>Cota Z deste PI apos o ajuste. Null quando geometricamente inavaliavel.</summary>
    public double? PosicaoZ { get; set; }

    /// <summary>Status consolidado do PI apos a execucao.</summary>
    public StatusAjustePI Status { get; set; }

    /// <summary>
    /// Verdadeiro quando o PI e um dos extremos internos (PI 1 ou PI n-2). Esses PIs
    /// dependem da elevacao dos pontos terminais (PI 0 e PI n-1), que por sua vez
    /// sao compartilhados com feature lines vizinhas - portanto violacoes neste ponto
    /// podem nao ter solucao local.
    /// </summary>
    public bool EhFronteiraInterna { get; set; }

    /// <summary>
    /// Indice do PI que foi movido quando este PI estava violando (ultimo alvo escolhido
    /// ao longo das passadas). Null quando este PI nunca violou. Pode ser igual ao proprio
    /// indice (ajuste direto) ou um vizinho (delegacao no modo Inteligente).
    /// </summary>
    public int? UltimoAlvoEscolhido { get; set; }

    /// <summary>
    /// Texto amigavel descrevendo qual estacao foi movida para resolver a violacao deste PI.
    /// Exemplos: "--" (nao violou), "ele proprio", "PI 7 (vizinho)".
    /// </summary>
    public string ResolvidoVia
    {
        get
        {
            if (UltimoAlvoEscolhido == null)
                return "--";
            if (UltimoAlvoEscolhido == IndicePI)
                return "ele proprio";
            int delta = UltimoAlvoEscolhido.Value - IndicePI;
            string direcao = delta < 0 ? "vizinho -" : "vizinho +";
            return $"PI {UltimoAlvoEscolhido} ({direcao}{System.Math.Abs(delta)})";
        }
    }

    /// <summary>Texto amigavel para a tabela, combinando status e flag de fronteira.</summary>
    public string StatusTexto => Status switch
    {
        StatusAjustePI.JaDentroDoLimite => "OK (inalterado)",
        StatusAjustePI.AjustadoComSucesso => "Ajustado OK",
        StatusAjustePI.ViolacaoRemanescente => EhFronteiraInterna
            ? "Viola (fronteira)"
            : "Viola (interno)",
        StatusAjustePI.FalhaGeometrica => "Falha geometrica",
        _ => Status.ToString()
    };

    /// <summary>Glifo visual compacto para coluna de icone (Unicode: check / x / alerta).</summary>
    public string Icone => Status switch
    {
        StatusAjustePI.JaDentroDoLimite => "\u2713",          // check
        StatusAjustePI.AjustadoComSucesso => "\u2713",        // check
        StatusAjustePI.ViolacaoRemanescente => "\u2717",      // ballot x
        StatusAjustePI.FalhaGeometrica => "\u26A0",           // warning
        _ => "?"
    };
}

/// <summary>
/// Resumo agregado de uma feature line para exibicao na tabela de resumo.
/// </summary>
public sealed class RelatorioFeatureLine
{
    /// <summary>Nome da feature line.</summary>
    public string Nome { get; set; } = string.Empty;

    /// <summary>Quantidade de PIs internos avaliados (indices 1..n-2 limitados por IndiceMaximo).</summary>
    public int TotalPIAvaliados { get; set; }

    /// <summary>Quantidade de violacoes detectadas no pre-scan.</summary>
    public int ViolacoesIniciais { get; set; }

    /// <summary>Quantidade total de violacoes remanescentes apos o ajuste.</summary>
    public int ViolacoesFinais { get; set; }

    /// <summary>
    /// Violacoes remanescentes em PI internos (nao fronteira). Este e o numero que
    /// realmente importa: violacoes na fronteira podem depender de FLs vizinhas.
    /// </summary>
    public int ViolacoesFinaisInternas { get; set; }

    /// <summary>Violacoes remanescentes em PI de fronteira (PI 1 ou PI n-2).</summary>
    public int ViolacoesFinaisFronteira { get; set; }

    /// <summary>Numero de estacoes distintas que tiveram a cota alterada.</summary>
    public int EstacoesAlteradas { get; set; }

    /// <summary>Total de iteracoes internas de ajuste aplicadas.</summary>
    public int Iteracoes { get; set; }

    /// <summary>Maior |dZ| aplicado em uma unica estacao desta FL.</summary>
    public double DeltaZMaximo { get; set; }

    /// <summary>Media dos |dZ| aplicados nas estacoes alteradas desta FL.</summary>
    public double DeltaZMedio { get; set; }

    /// <summary>Status consolidado da feature line apos o ajuste.</summary>
    public StatusRelatorioFL Status { get; set; }

    /// <summary>Texto amigavel para a tabela de resumo.</summary>
    public string StatusTexto => Status switch
    {
        StatusRelatorioFL.SemViolacoes => "Sem violacoes",
        StatusRelatorioFL.AjustadoComSucesso => "Ajustado com sucesso",
        StatusRelatorioFL.ViolacoesApenasEmFronteira => "Apenas fronteira",
        StatusRelatorioFL.ViolacoesInternasRemanescentes => "Violacoes internas",
        StatusRelatorioFL.Erro => "Erro",
        _ => Status.ToString()
    };

    /// <summary>Glifo visual compacto (check / alerta / x).</summary>
    public string Icone => Status switch
    {
        StatusRelatorioFL.SemViolacoes => "\u2713",
        StatusRelatorioFL.AjustadoComSucesso => "\u2713",
        StatusRelatorioFL.ViolacoesApenasEmFronteira => "\u26A0",
        StatusRelatorioFL.ViolacoesInternasRemanescentes => "\u2717",
        StatusRelatorioFL.Erro => "\u26A0",
        _ => "?"
    };
}

/// <summary>
/// Resumo da execucao do ajuste de deflexao para exibicao na linha de comando e na janela.
/// </summary>
public sealed class AjustarDeflexaoFeatureLinesResultado
{
    /// <summary>Quantidade total de feature lines encontradas apos filtros.</summary>
    public int TotalFeatureLinesConsideradas { get; set; }

    /// <summary>Quantidade de feature lines que receberam pelo menos um ajuste.</summary>
    public int TotalFeatureLinesAjustadas { get; set; }

    /// <summary>
    /// Quantidade total de PI points que violavam o limite de deflexao antes de qualquer ajuste,
    /// somando todas as feature lines consideradas.
    /// </summary>
    public int TotalViolacoesIniciais { get; set; }

    /// <summary>
    /// Quantidade total de PI points que continuam violando o limite apos a execucao.
    /// Ideal: zero. Valores &gt; 0 indicam pontos que nao convergiram ou falhas geometricas.
    /// </summary>
    public int TotalViolacoesFinais { get; set; }

    /// <summary>
    /// Quantidade de estacoes (PI points) distintas que tiveram a cota efetivamente alterada,
    /// somando todas as feature lines. Minimizar este numero e o objetivo do modo Inteligente.
    /// </summary>
    public int TotalEstacoesAlteradas { get; set; }

    /// <summary>
    /// Quantidade total de iteracoes internas de ajuste aplicadas (eventos de ajuste, nao
    /// necessariamente pontos distintos). Uma mesma estacao pode ser ajustada em varias iteracoes.
    /// </summary>
    public int TotalIteracoesAplicadas { get; set; }

    /// <summary>Quantidade de PI points ja dentro do limite (nao precisaram de ajuste).</summary>
    public int TotalPontosDentroDoLimite { get; set; }

    /// <summary>Quantidade de PI points que nao convergiram dentro do limite de iteracoes.</summary>
    public int TotalPontosNaoConvergidos { get; set; }

    /// <summary>Magnitude media dos |dZ| efetivamente aplicados (m).</summary>
    public double DeltaZMedio { get; set; }

    /// <summary>Maior |dZ| efetivamente aplicado em uma unica estacao (m).</summary>
    public double DeltaZMaximo { get; set; }

    /// <summary>
    /// Violacoes remanescentes em PI internos (nao fronteira). Soma sobre todas as FLs.
    /// E o indicador que realmente importa para avaliar sucesso.
    /// </summary>
    public int TotalViolacoesFinaisInternas { get; set; }

    /// <summary>
    /// Violacoes remanescentes em PI de fronteira (PI 1 ou PI n-2). Soma sobre todas as FLs.
    /// Essas violacoes podem ter origem em feature lines vizinhas e nao sao necessariamente "erro".
    /// </summary>
    public int TotalViolacoesFinaisFronteira { get; set; }

    /// <summary>Nomes das feature lines processadas com sucesso (pelo menos 1 ajuste).</summary>
    public List<string> NomesAjustadas { get; } = new();

    /// <summary>Resumo por feature line para a tabela de resumo na janela.</summary>
    public List<RelatorioFeatureLine> RelatorioFeatureLines { get; } = new();

    /// <summary>Detalhamento por PI para a tabela de detalhes na janela.</summary>
    public List<RelatorioPontoPI> RelatorioPontos { get; } = new();

    /// <summary>Logs detalhados de processamento por feature line e ponto.</summary>
    public List<string> Logs { get; } = new();

    /// <summary>Mensagens de erro ou de validacao acumuladas na execucao.</summary>
    public List<string> MensagensErro { get; } = new();
}
