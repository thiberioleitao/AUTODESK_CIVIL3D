using System;
using System.Collections.Generic;
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
/// Janela modeless para aplicar um Alignment Label Set Style a todos os
/// alinhamentos do desenho. Permanece aberta apos execucao para facilitar
/// repeticoes do comando com parametros diferentes.
/// </summary>
public partial class MudarLabelSetAlinhamentosWindow : Window
{
    /// <summary>
    /// Evento disparado quando o usuario clica em Executar e a validacao passa.
    /// </summary>
    public event EventHandler<MudarLabelSetAlinhamentosRequest>? ConfirmarClicado;

    public MudarLabelSetAlinhamentosWindow(IReadOnlyList<string> labelSetStyles)
    {
        InitializeComponent();

        ConfigurarComboBoxFiltravel(LabelSetComboBox, labelSetStyles, labelSetStyles.FirstOrDefault());
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

    public string NomeLabelSetStyle => LerValor(LabelSetComboBox);

    public bool ApagarExistentes => ApagarExistentesCheckBox.IsChecked == true;

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NomeLabelSetStyle))
        {
            MessageBox.Show(
                this,
                "Selecione ou digite o nome do Alignment Label Set Style antes de executar.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var request = new MudarLabelSetAlinhamentosRequest
        {
            NomeLabelSetStyle = NomeLabelSetStyle,
            ApagarExistentes = ApagarExistentes
        };

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
}
