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
/// Janela modeless da rotina "CRIAR REGIOES A PARTIR DE CORREDORES".
/// Permanece aberta apos execucao para facilitar ajuste de parametros.
/// </summary>
public partial class CriarRegioesCorredoresWindow : Window
{
    private static readonly CultureInfo s_cultura = CultureInfo.InvariantCulture;

    public event EventHandler<CriarRegioesCorredoresRequest>? ConfirmarClicado;

    public CriarRegioesCorredoresWindow(IReadOnlyList<string> assemblies)
    {
        InitializeComponent();

        ConfigurarComboBoxFiltravel(AssemblyComboBox, assemblies, assemblies.FirstOrDefault());

        // Valores padrao (mesmos do Dynamo original).
        AnguloDirecaoTextBox.Text = "60";
        VariacaoDeclividadeTextBox.Text = "0.05";
        EspDeclividadeTextBox.Text = "50";
        EspCruzamentosTextBox.Text = "50";

        FreqTangenteTextBox.Text = "1";
        IncCurvaHTextBox.Text = "0.5";
        MidOrdTextBox.Text = "1";
        FreqEspiralTextBox.Text = "1";

        FreqCurvaVTextBox.Text = "1";

        IncCurvaOffsetTextBox.Text = "1";
        MidOrdOffsetTextBox.Text = "1";
    }

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
            // Nao fatal.
        }
    }

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (!TentarLerParametros(out CriarRegioesCorredoresRequest? request) || request == null)
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

    private bool TentarLerParametros(out CriarRegioesCorredoresRequest? request)
    {
        request = null;

        string nomeAssembly = LerValor(AssemblyComboBox);
        if (string.IsNullOrWhiteSpace(nomeAssembly))
        {
            MessageBox.Show(this, "Selecione ou digite o nome do Assembly.", "ZagoCivil3D",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (!TentarLerDouble(AnguloDirecaoTextBox.Text, "Limite angulo de direcao", true, out double anguloDir)) return false;
        if (!TentarLerDouble(VariacaoDeclividadeTextBox.Text, "Variacao de declividade", true, out double varDecl)) return false;
        if (!TentarLerDouble(EspDeclividadeTextBox.Text, "Espacamento declividade", true, out double espDecl)) return false;
        if (!TentarLerDouble(EspCruzamentosTextBox.Text, "Espacamento cruzamentos", true, out double espCruz)) return false;

        if (!TentarLerDouble(FreqTangenteTextBox.Text, "Frequencia tangente", true, out double freqTan)) return false;
        if (!TentarLerDouble(IncCurvaHTextBox.Text, "Incremento curva horizontal", true, out double incCurvaH)) return false;
        if (!TentarLerDouble(MidOrdTextBox.Text, "Mid-ordinate horizontal", true, out double midOrd)) return false;
        if (!TentarLerDouble(FreqEspiralTextBox.Text, "Frequencia espiral", true, out double freqEsp)) return false;
        if (!TentarLerDouble(FreqCurvaVTextBox.Text, "Frequencia curva vertical", true, out double freqCurvaV)) return false;

        if (!TentarLerDouble(IncCurvaOffsetTextBox.Text, "Incremento curva offset", true, out double incCurvaOff)) return false;
        if (!TentarLerDouble(MidOrdOffsetTextBox.Text, "Mid-ordinate offset", true, out double midOrdOff)) return false;

        request = new CriarRegioesCorredoresRequest
        {
            NomeAssembly = nomeAssembly,
            LimiteAnguloDirecao = anguloDir,
            LimiteMudancaDeclividade = varDecl,
            EspacamentoMinimoDeclividade = espDecl,
            EspacamentoMinimoCruzamentos = espCruz,
            FrequenciaTangente = freqTan,
            IncrementoCurvaHorizontal = incCurvaH,
            DistanciaMidOrdinate = midOrd,
            FrequenciaEspiral = freqEsp,
            HorizontalEmGeometryPoints = HorizGeomPointsCheckBox.IsChecked == true,
            EmPontosSuperelevacao = SuperelevPointsCheckBox.IsChecked == true,
            FrequenciaCurvaVertical = freqCurvaV,
            VerticalEmGeometryPoints = VertGeomPointsCheckBox.IsChecked == true,
            VerticalEmHighLowPoints = HighLowPointsCheckBox.IsChecked == true,
            OffsetEmGeometryPoints = OffsetGeomPointsCheckBox.IsChecked == true,
            OffsetAdjacenteStartEnd = OffsetAdjStartEndCheckBox.IsChecked == true,
            IncrementoCurvaOffset = incCurvaOff,
            DistanciaMidOrdinateOffset = midOrdOff,
        };
        return true;
    }

    private bool TentarLerDouble(string? texto, string nomeCampo, bool positivo, out double valor)
    {
        valor = 0;
        string normalizado = (texto ?? string.Empty).Trim().Replace(',', '.');
        if (!double.TryParse(normalizado, NumberStyles.Float, s_cultura, out valor))
        {
            MessageBox.Show(this, $"O campo '{nomeCampo}' deve ser um numero valido.",
                "ZagoCivil3D", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        if (positivo && valor < 0)
        {
            MessageBox.Show(this, $"O campo '{nomeCampo}' nao pode ser negativo.",
                "ZagoCivil3D", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        return true;
    }

    private static string LerValor(ComboBox combo)
    {
        if (combo.SelectedItem is string itemSelecionado && !string.IsNullOrWhiteSpace(itemSelecionado))
            return itemSelecionado.Trim();
        return combo.Text?.Trim() ?? string.Empty;
    }

    /// <summary>
    /// Configura ComboBox editavel com filtro em tempo real, trocando
    /// diretamente o ItemsSource para preservar texto e cursor.
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
            try { combo.SelectedItem = inicial; }
            finally { supressor = false; }
        }

        combo.AddHandler(
            TextBoxBase.TextChangedEvent,
            new TextChangedEventHandler((s, e) =>
            {
                if (supressor) return;
                if (e.OriginalSource is not TextBox caixaTexto) return;

                string texto = caixaTexto.Text ?? string.Empty;
                int posicaoCursor = caixaTexto.CaretIndex;

                List<string> filtrados = string.IsNullOrEmpty(texto)
                    ? todos
                    : todos.Where(i => i.IndexOf(texto, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

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
