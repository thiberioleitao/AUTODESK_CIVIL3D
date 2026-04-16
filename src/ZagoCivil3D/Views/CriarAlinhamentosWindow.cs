using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Media;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Views;

/// <summary>
/// Tela modeless de coleta de parâmetros para criação de alignments
/// a partir de polilinhas de um layer, sem ordenação específica.
/// ComboBoxes são editáveis com filtro em tempo real para facilitar seleção
/// em listas longas de layers/estilos.
/// </summary>
public partial class CriarAlinhamentosWindow : Window
{
    /// <summary>
    /// Disparado quando o usuário confirma a entrada de dados válida.
    /// </summary>
    public event EventHandler<CriarAlinhamentosRequest>? ConfirmarClicado;

    public CriarAlinhamentosWindow(
        IReadOnlyList<string> camadas,
        IReadOnlyList<string> estilos,
        IReadOnlyList<string> conjuntosRotulos)
    {
        InitializeComponent();

        ConfigurarComboBoxFiltravel(LayerComboBox, camadas, camadas.FirstOrDefault());
        ConfigurarComboBoxFiltravel(StyleComboBox, estilos, estilos.FirstOrDefault());
        ConfigurarComboBoxFiltravel(LabelSetComboBox, conjuntosRotulos, conjuntosRotulos.FirstOrDefault());

        PrefixoTextBox.Text = "D";
        NumeroInicialTextBox.Text = "1";
        ApagarPolilinhasCheckBox.IsChecked = false;
    }

    /// <summary>
    /// Garante que a janela modeless seja parented ao AutoCAD e receba foco
    /// de teclado corretamente. Sem isso, WPF modeless em hosts como o Civil 3D
    /// pode ignorar eventos de tecla em TextBox/ComboBox editáveis.
    /// </summary>
    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        try
        {
            var helper = new WindowInteropHelper(this);
            IntPtr mainHwnd = Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.Handle;
            if (mainHwnd != IntPtr.Zero && helper.Owner == IntPtr.Zero)
                helper.Owner = mainHwnd;
        }
        catch
        {
            // Não-fatal: a janela ainda funciona sem o parent explícito,
            // mas com possíveis problemas de foco.
        }
    }

    /// <summary>
    /// Configura o ComboBox como editável com filtro em tempo real.
    /// Em vez de usar CollectionView.Filter (que sincroniza SelectedItem e
    /// acaba revertendo o texto digitado), trocamos diretamente o ItemsSource
    /// por uma lista filtrada e preservamos explicitamente Text e posição do
    /// cursor. Isso garante que Backspace/Delete e digitação funcionem de forma
    /// natural, sem "volta" para o item anterior.
    /// </summary>
    private static void ConfigurarComboBoxFiltravel(
        ComboBox combo,
        IReadOnlyList<string> itens,
        string? inicial)
    {
        var todos = itens.ToList();
        combo.ItemsSource = todos;

        bool supressor = false;

        if (!string.IsNullOrEmpty(inicial))
        {
            supressor = true;
            try
            {
                combo.SelectedItem = inicial;
            }
            finally
            {
                supressor = false;
            }
        }

        combo.AddHandler(
            TextBoxBase.TextChangedEvent,
            new TextChangedEventHandler((s, e) =>
            {
                if (supressor)
                    return;

                if (e.OriginalSource is not TextBox caixaTexto)
                    return;

                string texto = caixaTexto.Text ?? string.Empty;
                int posicaoCursor = caixaTexto.CaretIndex;

                List<string> filtrados = string.IsNullOrEmpty(texto)
                    ? todos
                    : todos
                        .Where(i => i.IndexOf(texto, StringComparison.OrdinalIgnoreCase) >= 0)
                        .ToList();

                supressor = true;
                try
                {
                    combo.SelectedItem = null;
                    combo.ItemsSource = filtrados;

                    if (caixaTexto.Text != texto)
                        caixaTexto.Text = texto;
                    if (caixaTexto.CaretIndex != posicaoCursor)
                        caixaTexto.CaretIndex = posicaoCursor;
                }
                finally
                {
                    supressor = false;
                }

                if (!combo.IsDropDownOpen && caixaTexto.IsKeyboardFocusWithin)
                    combo.IsDropDownOpen = true;
            }));
    }

    public string NomeCamadaOrigem => LerValor(LayerComboBox);
    public string NomeEstiloAlinhamento => LerValor(StyleComboBox);
    public string NomeConjuntoRotulosAlinhamento => LerValor(LabelSetComboBox);

    private static string LerValor(ComboBox combo)
    {
        if (combo.SelectedItem is string itemSelecionado && !string.IsNullOrWhiteSpace(itemSelecionado))
            return itemSelecionado.Trim();
        return combo.Text?.Trim() ?? string.Empty;
    }

    public string Prefixo => PrefixoTextBox.Text?.Trim() ?? string.Empty;
    public bool ApagarPolilinhasOriginais => ApagarPolilinhasCheckBox.IsChecked == true;
    public int NumeroInicial => int.Parse(NumeroInicialTextBox.Text.Trim());

    /// <summary>
    /// Exibe uma mensagem de status dentro da própria janela modeless.
    /// </summary>
    public void DefinirStatus(string mensagem, bool sucesso)
    {
        StatusBorder.Visibility = Visibility.Visible;
        StatusTextBlock.Text = mensagem;
        StatusBorder.Background = sucesso
            ? new SolidColorBrush(Color.FromRgb(0xEE, 0xF7, 0xEE))
            : new SolidColorBrush(Color.FromRgb(0xFD, 0xEC, 0xEC));
        StatusBorder.BorderBrush = sucesso
            ? new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50))
            : new SolidColorBrush(Color.FromRgb(0xE5, 0x3E, 0x3E));
        StatusTextBlock.Foreground = sucesso
            ? new SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0x32))
            : new SolidColorBrush(Color.FromRgb(0xA7, 0x1D, 0x2A));
    }

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NomeCamadaOrigem)
            || string.IsNullOrWhiteSpace(NomeEstiloAlinhamento)
            || string.IsNullOrWhiteSpace(NomeConjuntoRotulosAlinhamento)
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

        var requisicao = new CriarAlinhamentosRequest
        {
            Prefixo = Prefixo,
            NomeCamadaOrigem = NomeCamadaOrigem,
            NomeEstiloAlinhamento = NomeEstiloAlinhamento,
            NomeConjuntoRotulosAlinhamento = NomeConjuntoRotulosAlinhamento,
            NumeroInicial = NumeroInicial,
            ApagarPolilinhasOriginais = ApagarPolilinhasOriginais
        };

        ConfirmarButton.IsEnabled = false;
        CancelarButton.IsEnabled = false;
        try
        {
            ConfirmarClicado?.Invoke(this, requisicao);
        }
        finally
        {
            ConfirmarButton.IsEnabled = true;
            CancelarButton.IsEnabled = true;
        }
    }
}
