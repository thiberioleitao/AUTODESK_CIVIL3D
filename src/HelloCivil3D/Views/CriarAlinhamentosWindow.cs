using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace HelloCivil3D.Views;

/// <summary>
/// Tela de coleta de parâmetros de criação de alignments.
/// Expõe propriedades já normalizadas para simplificar o comando chamador.
/// </summary>
public partial class CriarAlinhamentosWindow : Window
{
    public CriarAlinhamentosWindow(
		IReadOnlyList<string> layers,
		IReadOnlyList<string> sites,
		IReadOnlyList<string> styles,
		IReadOnlyList<string> labelSets)
	{
		InitializeComponent();

		LayerComboBox.ItemsSource = layers;
		SiteComboBox.ItemsSource = sites;
		StyleComboBox.ItemsSource = styles;
		LabelSetComboBox.ItemsSource = labelSets;

		LayerComboBox.SelectedItem = layers.FirstOrDefault();
		SiteComboBox.SelectedItem = sites.FirstOrDefault();
		StyleComboBox.SelectedItem = styles.FirstOrDefault();
		LabelSetComboBox.SelectedItem = labelSets.FirstOrDefault();

		PrefixoTextBox.Text = "D";
		NumeroInicialTextBox.Text = "1";
		IncrementoTextBox.Text = "1";
		ApagarPolilinhasCheckBox.IsChecked = false;
	}

	public string SourceLayerName => LayerComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;
	public string SiteName => SiteComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;
	public string AlignmentStyleName => StyleComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;
	public string AlignmentLabelSetName => LabelSetComboBox.SelectedItem?.ToString()?.Trim() ?? string.Empty;
	public string Prefixo => PrefixoTextBox.Text?.Trim() ?? string.Empty;
	public bool ApagarPolilinhasOriginais => ApagarPolilinhasCheckBox.IsChecked == true;

	public int NumeroInicial => int.Parse(NumeroInicialTextBox.Text.Trim());
	public int Incremento => int.Parse(IncrementoTextBox.Text.Trim());

    private void CancelarButton_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
		Close();
	}

    private void ConfirmarButton_Click(object sender, RoutedEventArgs e)
	{
      // A validação ocorre na borda da UI para impedir requests inválidos
		// chegarem ao serviço de domínio.
		if (string.IsNullOrWhiteSpace(SourceLayerName)
			|| string.IsNullOrWhiteSpace(SiteName)
			|| string.IsNullOrWhiteSpace(AlignmentStyleName)
			|| string.IsNullOrWhiteSpace(AlignmentLabelSetName)
			|| string.IsNullOrWhiteSpace(Prefixo))
		{
			MessageBox.Show(
				this,
				"Preencha os campos obrigatórios antes de confirmar.",
				"HelloCivil3D",
				MessageBoxButton.OK,
				MessageBoxImage.Warning);
			return;
		}

		if (!int.TryParse(NumeroInicialTextBox.Text?.Trim(), out int numeroInicial) || numeroInicial <= 0)
		{
			MessageBox.Show(
				this,
				"Informe um Número inicial válido maior que zero.",
				"HelloCivil3D",
				MessageBoxButton.OK,
				MessageBoxImage.Warning);
			return;
		}

		if (!int.TryParse(IncrementoTextBox.Text?.Trim(), out int incremento) || incremento <= 0)
		{
			MessageBox.Show(
				this,
				"Informe um Incremento válido maior que zero.",
				"HelloCivil3D",
				MessageBoxButton.OK,
				MessageBoxImage.Warning);
			return;
		}

		DialogResult = true;
		Close();
	}
}
