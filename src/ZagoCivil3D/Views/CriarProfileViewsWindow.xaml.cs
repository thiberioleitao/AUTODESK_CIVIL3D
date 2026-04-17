using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Media;
using ZagoCivil3D.Models;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ZagoCivil3D.Views;

/// <summary>
/// Janela modeless para criacao automatica de profile views a partir dos
/// alinhamentos do desenho. Permanece aberta apos execucao para facilitar
/// repeticoes do comando com parametros diferentes.
/// </summary>
public partial class CriarProfileViewsWindow : Window
{
    private static readonly CultureInfo s_cultura = CultureInfo.InvariantCulture;

    /// <summary>
    /// Evento disparado quando o usuario clica em Executar e a validacao passa.
    /// </summary>
    public event EventHandler<CriarProfileViewsRequest>? ConfirmarClicado;

    public CriarProfileViewsWindow(
        IReadOnlyList<string> estilosProfileView,
        IReadOnlyList<string> conjuntosBands)
    {
        InitializeComponent();

        ConfigurarComboBoxFiltravel(EstiloProfileViewComboBox, estilosProfileView, estilosProfileView.FirstOrDefault());
        ConfigurarComboBoxFiltravel(ConjuntoBandsComboBox, conjuntosBands, conjuntosBands.FirstOrDefault());

        // Valores padrao (extraidos do Dynamo 1.3 como referencia).
        CoordenadaXTextBox.Text = "0";
        CoordenadaYTextBox.Text = "0";
        OffsetAdicionalTextBox.Text = "50";
        SufixoProfileViewTextBox.Text = "_PV";
    }

    /// <summary>
    /// Define a janela do AutoCAD como owner para garantir foco e teclado corretos.
    /// </summary>
    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        try
        {
            var helper = new WindowInteropHelper(this);
            IntPtr mainHwnd = AplicacaoAutoCad.MainWindow.Handle;
            if (mainHwnd != IntPtr.Zero && helper.Owner == IntPtr.Zero)
                helper.Owner = mainHwnd;
        }
        catch
        {
            // Nao fatal: janela funciona sem owner explicito
        }
    }

    public string NomeEstiloProfileView => LerValor(EstiloProfileViewComboBox);

    public string NomeConjuntoBands => LerValor(ConjuntoBandsComboBox);

    public string SufixoProfileView => SufixoProfileViewTextBox.Text?.Trim() ?? "_PV";

    public bool SubstituirExistentes => SubstituirExistentesCheckBox.IsChecked == true;

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NomeEstiloProfileView)
            || string.IsNullOrWhiteSpace(NomeConjuntoBands))
        {
            MessageBox.Show(
                this,
                "Selecione o estilo do profile view e o conjunto de bands antes de executar.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(SufixoProfileView))
        {
            MessageBox.Show(
                this,
                "Informe um sufixo para o nome do profile view.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!TentarLerParametros(out CriarProfileViewsRequest? request) || request == null)
            return;

        ConfirmarButton.IsEnabled = false;
        CancelarButton.IsEnabled = false;
        try
        {
            ConfirmarClicado?.Invoke(this, request);
        }
        finally
        {
            ConfirmarButton.IsEnabled = true;
            CancelarButton.IsEnabled = true;
        }
    }

    /// <summary>
    /// Exibe uma mensagem de status na barra inferior da janela.
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

    private bool TentarLerParametros(out CriarProfileViewsRequest? request)
    {
        request = null;

        if (!TentarLerDouble(CoordenadaXTextBox.Text, "Coordenada X", permitirNegativo: true, out double coordenadaX))
            return false;

        if (!TentarLerDouble(CoordenadaYTextBox.Text, "Coordenada Y", permitirNegativo: true, out double coordenadaY))
            return false;

        if (!TentarLerDouble(OffsetAdicionalTextBox.Text, "Offset adicional", permitirNegativo: false, out double offset))
            return false;

        request = new CriarProfileViewsRequest
        {
            CoordenadaX = coordenadaX,
            CoordenadaY = coordenadaY,
            OffsetAdicional = offset,
            NomeEstiloProfileView = NomeEstiloProfileView,
            NomeConjuntoBands = NomeConjuntoBands,
            SufixoProfileView = SufixoProfileView,
            SubstituirExistentes = SubstituirExistentes
        };

        return true;
    }

    /// <summary>
    /// Le o valor de um ComboBox editavel, priorizando o item selecionado.
    /// </summary>
    private static string LerValor(ComboBox combo)
    {
        if (combo.SelectedItem is string itemSelecionado && !string.IsNullOrWhiteSpace(itemSelecionado))
            return itemSelecionado.Trim();

        return combo.Text?.Trim() ?? string.Empty;
    }

    /// <summary>
    /// Configura um ComboBox com filtro por texto digitado.
    /// Troca diretamente o ItemsSource preservando texto e cursor.
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

    private bool TentarLerDouble(
        string? texto,
        string nomeCampo,
        bool permitirNegativo,
        out double valor)
    {
        valor = 0;
        string normalizado = (texto ?? string.Empty).Trim().Replace(',', '.');

        if (!double.TryParse(normalizado, NumberStyles.Float, s_cultura, out valor))
        {
            MessageBox.Show(
                this,
                $"Informe um valor numerico valido para '{nomeCampo}'.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        if (!permitirNegativo && valor < 0)
        {
            MessageBox.Show(
                this,
                $"O campo '{nomeCampo}' nao pode ser negativo.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        return true;
    }
}
