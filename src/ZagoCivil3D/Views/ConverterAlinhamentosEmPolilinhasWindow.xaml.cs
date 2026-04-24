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
/// Janela modeless do comando que converte alinhamentos em polilinhas 2D.
/// Permanece aberta apos a execucao para permitir novas execucoes com
/// parametros diferentes sem reabrir o comando.
/// </summary>
public partial class ConverterAlinhamentosEmPolilinhasWindow : Window
{
    private static readonly CultureInfo s_cultura = CultureInfo.InvariantCulture;

    /// <summary>
    /// Evento disparado quando o usuario clica em Executar e a validacao passa.
    /// </summary>
    public event EventHandler<ConverterAlinhamentosEmPolilinhasRequest>? ConfirmarClicado;

    public ConverterAlinhamentosEmPolilinhasWindow(IReadOnlyList<string> layersDesenho)
    {
        InitializeComponent();

        ConfigurarComboBoxLayer(LayerComboBox, layersDesenho, "C-ROAD-POLY");

        ElevationTextBox.Text = "0";
        PassoDiscretizacaoTextBox.Text = "1.0";
    }

    /// <summary>
    /// Define a janela do AutoCAD como owner para garantir foco de teclado
    /// correto nos campos editaveis (fix obrigatorio para janelas modeless
    /// no Civil 3D).
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
            // Nao fatal: janela funciona sem owner explicito.
        }
    }

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (!TentarLerParametros(out ConverterAlinhamentosEmPolilinhasRequest? request) || request == null)
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

    /// <summary>
    /// Le e valida os parametros informados no formulario.
    /// </summary>
    private bool TentarLerParametros(out ConverterAlinhamentosEmPolilinhasRequest? request)
    {
        request = null;

        if (!TentarLerDouble(ElevationTextBox.Text, "Elevation (Z)", permitirNegativo: true, out double elevation))
            return false;

        if (!TentarLerDouble(PassoDiscretizacaoTextBox.Text, "Passo discretizacao", permitirNegativo: false, out double passo))
            return false;

        if (passo <= 0)
        {
            AvisarValor("Passo discretizacao", "deve ser maior que zero.");
            return false;
        }

        request = new ConverterAlinhamentosEmPolilinhasRequest
        {
            FiltroNome = FiltroNomeTextBox.Text?.Trim() ?? string.Empty,
            NomeLayer = LerValor(LayerComboBox),
            Elevation = elevation,
            PreservarArcos = PreservarArcosCheckBox.IsChecked == true,
            PassoDiscretizacaoEspirais = passo,
            ApagarAlinhamentosOriginais = ApagarAlinhamentosCheckBox.IsChecked == true,
            DryRun = DryRunCheckBox.IsChecked == true,
        };

        return true;
    }

    private void AvisarValor(string campo, string mensagem)
    {
        MessageBox.Show(
            this,
            $"O campo '{campo}' {mensagem}",
            "ZagoCivil3D",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);
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
            AvisarValor(nomeCampo, "deve ser um numero valido.");
            return false;
        }

        if (!permitirNegativo && valor < 0)
        {
            AvisarValor(nomeCampo, "nao pode ser negativo.");
            return false;
        }

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
    /// Alimenta um ComboBox de layer com as layers existentes no desenho,
    /// garantindo que o nome padrao tambem apareca na lista.
    /// </summary>
    private static void ConfigurarComboBoxLayer(
        ComboBox combo,
        IReadOnlyList<string> layersDesenho,
        string nomePadrao)
    {
        var lista = new List<string>(layersDesenho);
        if (!lista.Any(n => string.Equals(n, nomePadrao, StringComparison.OrdinalIgnoreCase)))
            lista.Insert(0, nomePadrao);

        string selecionado = lista.FirstOrDefault(
            n => string.Equals(n, nomePadrao, StringComparison.OrdinalIgnoreCase)) ?? nomePadrao;

        ConfigurarComboBoxFiltravel(combo, lista, selecionado);
    }

    /// <summary>
    /// Configura um ComboBox editavel com filtro em tempo real por substring.
    /// Swap direto do ItemsSource preservando texto e cursor.
    /// </summary>
    private static void ConfigurarComboBoxFiltravel(
        ComboBox combo,
        IReadOnlyList<string> itens,
        string? inicial)
    {
        var todos = itens.ToList();
        combo.ItemsSource = todos;

        bool supressor = false;

        if (inicial != null)
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
}
