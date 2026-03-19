using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace ZagoCivil3D.Views;

/// <summary>
/// Tela de coleta de parâmetros de criação de alignments.
/// Expõe propriedades já normalizadas para simplificar o comando chamador.
/// </summary>
public partial class CriarAlinhamentosWindow : Window
{
    public CriarAlinhamentosWindow(
       IReadOnlyList<string> camadas,
		IReadOnlyList<string> estilos,
		IReadOnlyList<string> conjuntosRotulos)
	{
		InitializeComponent();

        LayerComboBox.ItemsSource = camadas;
		StyleComboBox.ItemsSource = estilos;
		LabelSetComboBox.ItemsSource = conjuntosRotulos;

        LayerComboBox.SelectedItem = camadas.FirstOrDefault();
		StyleComboBox.SelectedItem = estilos.FirstOrDefault();
		LabelSetComboBox.SelectedItem = conjuntosRotulos.FirstOrDefault();

		PrefixoTextBox.Text = "D";
        ZonaTextBox.Text = "01";
		NumeroInicialTextBox.Text = "1";
		IncrementoTextBox.Text = "1";
		ApagarPolilinhasCheckBox.IsChecked = false;
	}

    public string NomeCamadaOrigem => LayerComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;
	public string NomeEstiloAlinhamento => StyleComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;
	public string NomeConjuntoRotulosAlinhamento => LabelSetComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;
	public string Prefixo => PrefixoTextBox.Text?.Trim() ?? string.Empty;
    public string IdentificadorZona => ZonaTextBox.Text?.Trim() ?? string.Empty;
	public bool ApagarPolilinhasOriginais => ApagarPolilinhasCheckBox.IsChecked == true;

	public int NumeroInicial => int.Parse(NumeroInicialTextBox.Text.Trim());
	public int Incremento => int.Parse(IncrementoTextBox.Text.Trim());

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
		Close();
	}

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
	{
      // A validação ocorre na borda da UI para impedir requests inválidos
		// chegarem ao serviço de domínio.
      if (string.IsNullOrWhiteSpace(NomeCamadaOrigem)
			|| string.IsNullOrWhiteSpace(NomeEstiloAlinhamento)
			|| string.IsNullOrWhiteSpace(NomeConjuntoRotulosAlinhamento)
          || string.IsNullOrWhiteSpace(IdentificadorZona)
			|| string.IsNullOrWhiteSpace(Prefixo))
		{
			MessageBox.Show(
				this,
				"Preencha os campos obrigatórios antes de confirmar.",
             "ZagoCivil3D",
				MessageBoxButton.OK,
				MessageBoxImage.Warning);
			return;
		}

		if (!int.TryParse(NumeroInicialTextBox.Text?.Trim(), out int numeroInicial) || numeroInicial <= 0)
		{
			MessageBox.Show(
				this,
				"Informe um Número inicial válido maior que zero.",
             "ZagoCivil3D",
				MessageBoxButton.OK,
				MessageBoxImage.Warning);
			return;
		}

		if (!int.TryParse(IncrementoTextBox.Text?.Trim(), out int incremento) || incremento <= 0)
		{
			MessageBox.Show(
				this,
				"Informe um Incremento válido maior que zero.",
             "ZagoCivil3D",
				MessageBoxButton.OK,
				MessageBoxImage.Warning);
			return;
		}

		DialogResult = true;
		Close();
	}
}
