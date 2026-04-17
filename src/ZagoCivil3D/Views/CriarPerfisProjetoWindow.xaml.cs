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
/// Janela modeless para criacao automatica de perfis de projeto.
/// Permanece aberta apos execucao para facilitar repeticoes do comando.
/// </summary>
public partial class CriarPerfisProjetoWindow : Window
{
    private static readonly CultureInfo s_cultura = CultureInfo.InvariantCulture;

    /// <summary>
    /// Evento disparado quando o usuario clica em Executar e a validacao passa.
    /// O comando se inscreve nesse evento para executar o processamento.
    /// </summary>
    public event EventHandler<CriarPerfisProjetoRequest>? ConfirmarClicado;

    public CriarPerfisProjetoWindow(
        IReadOnlyList<string> superficies,
        IReadOnlyList<string> estilosPerfil,
        IReadOnlyList<string> conjuntosRotulos,
        IReadOnlyList<string> camadas)
    {
        InitializeComponent();

        ConfigurarComboBoxFiltravel(SuperficieComboBox, superficies, superficies.FirstOrDefault());
        ConfigurarComboBoxFiltravel(EstiloPerfilComboBox, estilosPerfil, estilosPerfil.FirstOrDefault());
        ConfigurarComboBoxFiltravel(ConjuntoRotulosComboBox, conjuntosRotulos, conjuntosRotulos.FirstOrDefault());
        ConfigurarComboBoxFiltravel(CamadaComboBox, camadas, camadas.FirstOrDefault());

        EspacamentoEstacasTextBox.Text = "10";
        OffsetAbaixoTerrenoTextBox.Text = "0.100";
        DeclivedadeMinimaTextBox.Text = "0.002";
        LimiteVariacaoTextBox.Text = "0.050";
        SufixoPerfilTextBox.Text = "_PROJETO";
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

    public string NomeSuperficie => LerValor(SuperficieComboBox);
    public string NomeEstiloPerfil => LerValor(EstiloPerfilComboBox);
    public string NomeConjuntoRotulosPerfil => LerValor(ConjuntoRotulosComboBox);
    public string NomeCamada => LerValor(CamadaComboBox);
    public string SufixoPerfil => SufixoPerfilTextBox.Text?.Trim() ?? "_PROJETO";
    public bool SubstituirPerfisExistentes => SubstituirPerfisCheckBox.IsChecked == true;

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(NomeSuperficie)
            || string.IsNullOrWhiteSpace(NomeEstiloPerfil)
            || string.IsNullOrWhiteSpace(NomeConjuntoRotulosPerfil)
            || string.IsNullOrWhiteSpace(NomeCamada))
        {
            MessageBox.Show(
                this,
                "Preencha os campos obrigatorios antes de executar.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(SufixoPerfil))
        {
            MessageBox.Show(
                this,
                "Informe um sufixo para o nome do perfil.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!TentarLerParametros(out CriarPerfisProjetoRequest? request) || request == null)
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

    private bool TentarLerParametros(out CriarPerfisProjetoRequest? request)
    {
        request = null;

        if (!TentarLerDouble(EspacamentoEstacasTextBox.Text, "Espacamento entre estacoes", valorMinimo: 0.1, out double espacamento))
            return false;

        if (!TentarLerDouble(OffsetAbaixoTerrenoTextBox.Text, "Offset abaixo do terreno", valorMinimo: 0, out double offset))
            return false;

        if (!TentarLerDouble(DeclivedadeMinimaTextBox.Text, "Declividade minima", valorMinimo: 0, out double declivedade))
            return false;

        if (!TentarLerDouble(LimiteVariacaoTextBox.Text, "Limite de variacao de declividade", valorMinimo: 0, out double limiteVariacao))
            return false;

        request = new CriarPerfisProjetoRequest
        {
            NomeSuperficie = NomeSuperficie,
            NomeEstiloPerfil = NomeEstiloPerfil,
            NomeConjuntoRotulosPerfil = NomeConjuntoRotulosPerfil,
            NomeCamada = NomeCamada,
            EspacamentoEstacas = espacamento,
            OffsetAbaixoTerreno = offset,
            DeclivedadeMinima = declivedade,
            LimiteVariacaoDeclividade = limiteVariacao,
            SufixoPerfil = SufixoPerfil,
            SubstituirPerfisExistentes = SubstituirPerfisExistentes
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
        double valorMinimo,
        out double valor)
    {
        valor = 0;
        string normalizado = (texto ?? string.Empty).Trim().Replace(',', '.');

        if (!double.TryParse(normalizado, NumberStyles.Float, s_cultura, out valor)
            || valor < valorMinimo)
        {
            MessageBox.Show(
                this,
                $"Informe um valor valido para '{nomeCampo}' maior ou igual a {valorMinimo.ToString(s_cultura)}.",
                "ZagoCivil3D",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        return true;
    }
}
