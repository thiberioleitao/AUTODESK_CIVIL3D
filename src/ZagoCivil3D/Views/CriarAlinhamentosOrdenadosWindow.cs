using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Views;

/// <summary>
/// Tela modeless de coleta de parâmetros para criação de alignments ordenados:
/// horizontais Norte→Sul primeiro e verticais Oeste→Leste depois.
/// ComboBoxes são editáveis com filtro em tempo real para facilitar seleção
/// em listas longas de layers/estilos.
/// </summary>
public partial class CriarAlinhamentosOrdenadosWindow : Window
{
    /// <summary>
    /// Disparado quando o usuário confirma a entrada de dados válida.
    /// </summary>
    public event EventHandler<CriarAlinhamentosOrdenadosRequest>? ConfirmarClicado;

    public CriarAlinhamentosOrdenadosWindow(
        IReadOnlyList<string> camadas,
        IReadOnlyList<string> estilos,
        IReadOnlyList<string> conjuntosRotulos)
    {
        InitializeComponent();

        // Cada ComboBox recebe sua própria lista. Trocamos o ItemsSource
        // diretamente ao filtrar em vez de usar CollectionView.Filter, pois
        // Refresh() da view se sincroniza com SelectedItem e acaba revertendo
        // o texto digitado pelo usuário.
        ConfigurarComboBoxFiltravel(LayerHorizontaisComboBox, camadas, camadas.FirstOrDefault());
        ConfigurarComboBoxFiltravel(LayerVerticaisComboBox, camadas, camadas.FirstOrDefault());
        ConfigurarComboBoxFiltravel(StyleComboBox, estilos, estilos.FirstOrDefault());
        ConfigurarComboBoxFiltravel(LabelSetComboBox, conjuntosRotulos, conjuntosRotulos.FirstOrDefault());

        PrefixoTextBox.Text = "D";
        NumeroInicialTextBox.Text = "1";
        ApagarPolilinhasCheckBox.IsChecked = false;

        OrdemHorizontaisPrimeiroRadio.Checked += (_, _) => AtualizarResumoOrdem();
        OrdemVerticaisPrimeiroRadio.Checked += (_, _) => AtualizarResumoOrdem();
        AtualizarResumoOrdem();
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
    /// Atualiza o bloco de resumo no cabeçalho conforme a ordem selecionada.
    /// Mantém o usuário informado sem precisar rolar até os RadioButtons.
    /// </summary>
    private void AtualizarResumoOrdem()
    {
        OrdemResumoTextBlock.Inlines.Clear();

        var corAcento = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4));
        corAcento.Freeze();

        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.Run("Ordem de criação:")
        {
            FontWeight = FontWeights.SemiBold
        });
        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.LineBreak());

        (string primeiro, string segundo) = HorizontaisPrimeiro
            ? ("Horizontais (L↔O) — ordenadas do Norte para o Sul",
               "Verticais (N↕S) — ordenadas do Oeste para o Leste")
            : ("Verticais (N↕S) — ordenadas do Oeste para o Leste",
               "Horizontais (L↔O) — ordenadas do Norte para o Sul");

        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.Run("  1º  ")
        {
            FontWeight = FontWeights.Bold,
            Foreground = corAcento
        });
        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.Run(primeiro));
        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.LineBreak());

        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.Run("  2º  ")
        {
            FontWeight = FontWeights.Bold,
            Foreground = corAcento
        });
        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.Run(segundo));
        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.LineBreak());

        OrdemResumoTextBlock.Inlines.Add(new System.Windows.Documents.Run("Numeração contínua entre os dois grupos.")
        {
            FontStyle = FontStyles.Italic
        });
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
                    // Limpa a seleção antes de trocar o ItemsSource para evitar
                    // que o WPF re-sincronize Text ↔ SelectedItem e sobrescreva
                    // o que o usuário acabou de digitar.
                    combo.SelectedItem = null;
                    combo.ItemsSource = filtrados;

                    // ItemsSource change costuma limpar o Text; restauramos.
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

    // Com IsEditable=true, a propriedade Text reflete o item selecionado ou o
    // texto digitado. Lemos de Text para cobrir ambos os casos.
    public string NomeCamadaHorizontais => LerValor(LayerHorizontaisComboBox);
    public string NomeCamadaVerticais => LerValor(LayerVerticaisComboBox);
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
    public bool HorizontaisPrimeiro => OrdemHorizontaisPrimeiroRadio.IsChecked == true;

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
        // Validação de borda: impede request inválido chegar no serviço de domínio.
        if (string.IsNullOrWhiteSpace(NomeCamadaHorizontais)
            || string.IsNullOrWhiteSpace(NomeCamadaVerticais)
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

        if (string.Equals(NomeCamadaHorizontais, NomeCamadaVerticais, StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(
                this,
                "Os layers de polilinhas horizontais e verticais devem ser diferentes.",
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

        var requisicao = new CriarAlinhamentosOrdenadosRequest
        {
            Prefixo = Prefixo,
            NomeCamadaHorizontais = NomeCamadaHorizontais,
            NomeCamadaVerticais = NomeCamadaVerticais,
            NomeEstiloAlinhamento = NomeEstiloAlinhamento,
            NomeConjuntoRotulosAlinhamento = NomeConjuntoRotulosAlinhamento,
            NumeroInicial = NumeroInicial,
            ApagarPolilinhasOriginais = ApagarPolilinhasOriginais,
            HorizontaisPrimeiro = HorizontaisPrimeiro
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
