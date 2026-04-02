using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Views;

/// <summary>
/// Tela responsavel por coletar os parametros do fluxo de terraplenagem.
/// </summary>
public partial class TerraplenagemFeatureLinesWindow : Window
{
    /// <summary>
    /// Cultura fixa para padronizar a leitura e escrita de valores decimais.
    /// </summary>
    private static readonly CultureInfo s_cultura = CultureInfo.InvariantCulture;

    /// <summary>
    /// Inicializa a tela com as listas disponiveis no desenho atual.
    /// </summary>
    public TerraplenagemFeatureLinesWindow(
        IReadOnlyList<string> sites,
        IReadOnlyList<string> camadas,
        IReadOnlyList<string> superficies)
    {
        InitializeComponent();

        SiteComboBox.ItemsSource = sites;
        LayerComboBox.ItemsSource = camadas;
        SurfaceComboBox.ItemsSource = superficies;

        // Os presets abaixo sao apenas conveniencias baseadas no desenho de referencia.
        // Quando nao existirem, a tela cai para o primeiro item disponivel.
        SiteComboBox.SelectedItem = SelecionarValorPreferencial(sites, "TER - SUL");
        LayerComboBox.SelectedItem = SelecionarValorPreferencial(camadas, "ZAGO-SELECIONAR FEATURELINES");
        SurfaceComboBox.SelectedItem = SelecionarValorPreferencial(superficies, "0. Primitivo B");

        DeflexaoLimiteTextBox.Text = "0.0129";
        NumeroMaximoPontosTextBox.Text = "1000";
        PassoIncrementalTextBox.Text = "0.01";
        QuantidadePassadasTextBox.Text = "2";
        PercentualObjetivoTextBox.Text = "1";
        NumeroTentativasTextBox.Text = "100";
        ToleranciaBaixaTextBox.Text = "0.01";
        ToleranciaAltaTextBox.Text = "0.10";
        MaximoAjustesTextBox.Text = "1";
    }

    /// <summary>
    /// Nome do site selecionado na UI.
    /// </summary>
    public string NomeSite => SiteComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;

    /// <summary>
    /// Nome da layer que contem o poligono de filtro.
    /// </summary>
    public string NomeCamadaPoligono => LayerComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;

    /// <summary>
    /// Nome da superficie base selecionada.
    /// </summary>
    public string NomeSuperficieBase => SurfaceComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;

    /// <summary>
    /// Fecha a janela sem executar o fluxo.
    /// </summary>
    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    /// <summary>
    /// Valida os campos digitados e libera a execucao do comando.
    /// </summary>
    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NomeSite)
            || string.IsNullOrWhiteSpace(NomeCamadaPoligono)
            || string.IsNullOrWhiteSpace(NomeSuperficieBase))
        {
            MessageBox.Show(
                this,
                "Preencha os campos obrigatorios antes de executar.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!TentarLerParametros(out _))
            return;

        DialogResult = true;
        Close();
    }

    /// <summary>
    /// Converte os valores da tela em um objeto fortemente tipado.
    /// </summary>
    public TerraplenagemFeatureLinesRequest CriarRequest()
    {
        if (!TentarLerParametros(out TerraplenagemFeatureLinesRequest? request) || request == null)
            throw new InvalidOperationException("Nao foi possivel converter os parametros da tela.");

        return request;
    }

    /// <summary>
    /// Faz a validacao centralizada de todos os campos numericos da UI.
    /// </summary>
    private bool TentarLerParametros(out TerraplenagemFeatureLinesRequest? request)
    {
        request = null;

        if (!TentarLerDouble(DeflexaoLimiteTextBox.Text, "Deflexao limite", valorMinimo: 0, out double deflexaoLimite))
            return false;

        if (!TentarLerInteiro(NumeroMaximoPontosTextBox.Text, "Numero maximo de pontos por FL", valorMinimo: 2, out int numeroMaximoPontos))
            return false;

        if (!TentarLerDouble(PassoIncrementalTextBox.Text, "Passo incremental do ajuste de cota", valorMinimo: 0.000001, out double passoIncremental))
            return false;

        if (!TentarLerInteiro(QuantidadePassadasTextBox.Text, "Quantidade de passadas globais", valorMinimo: 1, out int quantidadePassadas))
            return false;

        if (!TentarLerDouble(PercentualObjetivoTextBox.Text, "Percentual objetivo", valorMinimo: 0, valorMaximo: 1, out double percentualObjetivo))
            return false;

        if (!TentarLerInteiro(NumeroTentativasTextBox.Text, "Numero de tentativas por ponto", valorMinimo: 1, out int numeroTentativas))
            return false;

        if (!TentarLerDouble(ToleranciaBaixaTextBox.Text, "Tolerancia baixa da superficie", valorMinimo: 0, out double toleranciaBaixa))
            return false;

        if (!TentarLerDouble(ToleranciaAltaTextBox.Text, "Tolerancia alta da superficie", valorMinimo: 0, out double toleranciaAlta))
            return false;

        if (!TentarLerInteiro(MaximoAjustesTextBox.Text, "Maximo de ajustes por feature line", valorMinimo: 1, out int maximoAjustes))
            return false;

        request = new TerraplenagemFeatureLinesRequest
        {
            NomeSite = NomeSite,
            NomeCamadaPoligono = NomeCamadaPoligono,
            NomeSuperficieBase = NomeSuperficieBase,
            DeflexaoLimite = deflexaoLimite,
            NumeroMaximoPontosPorFeatureLine = numeroMaximoPontos,
            PassoIncrementalAjusteCota = passoIncremental,
            QuantidadePassadasGlobais = quantidadePassadas,
            PercentualObjetivo = percentualObjetivo,
            NumeroTentativasPorPonto = numeroTentativas,
            ToleranciaBaixaSuperficie = toleranciaBaixa,
            ToleranciaAltaSuperficie = toleranciaAlta,
            MaximoAjustesPorFeatureLine = maximoAjustes
        };

        return true;
    }

    /// <summary>
    /// Tenta aplicar um preset conhecido sem tornar a tela dependente dele.
    /// </summary>
    private static string? SelecionarValorPreferencial(
        IReadOnlyList<string> itens,
        string valorPreferencial)
    {
        return itens.FirstOrDefault(x => string.Equals(x, valorPreferencial, StringComparison.OrdinalIgnoreCase))
            ?? itens.FirstOrDefault();
    }

    /// <summary>
    /// Le um valor inteiro garantindo a faixa minima esperada.
    /// </summary>
    private bool TentarLerInteiro(string? texto, string nomeCampo, int valorMinimo, out int valor)
    {
        valor = 0;

        if (!int.TryParse(texto?.Trim(), NumberStyles.Integer, s_cultura, out valor) || valor < valorMinimo)
        {
            MessageBox.Show(
                this,
                $"Informe um valor valido para '{nomeCampo}' maior ou igual a {valorMinimo}.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    /// <summary>
    /// Le um valor decimal permitindo ponto ou virgula como separador.
    /// </summary>
    private bool TentarLerDouble(
        string? texto,
        string nomeCampo,
        double valorMinimo,
        out double valor)
    {
        return TentarLerDouble(texto, nomeCampo, valorMinimo, null, out valor);
    }

    /// <summary>
    /// Le um valor decimal validando limites minimo e maximo quando necessario.
    /// </summary>
    private bool TentarLerDouble(
        string? texto,
        string nomeCampo,
        double valorMinimo,
        double? valorMaximo,
        out double valor)
    {
        valor = 0;
        string normalizado = (texto ?? string.Empty).Trim().Replace(',', '.');

        if (!double.TryParse(normalizado, NumberStyles.Float, s_cultura, out valor)
            || valor < valorMinimo
            || (valorMaximo.HasValue && valor > valorMaximo.Value))
        {
            string faixa = valorMaximo.HasValue
                ? $"entre {valorMinimo.ToString(s_cultura)} e {valorMaximo.Value.ToString(s_cultura)}"
                : $"maior ou igual a {valorMinimo.ToString(s_cultura)}";

            MessageBox.Show(
                this,
                $"Informe um valor valido para '{nomeCampo}' {faixa}.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        return true;
    }
}
