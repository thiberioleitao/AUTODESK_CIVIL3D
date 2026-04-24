using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
/// Janela modeless para exportar o CSV de bacias/subbacias. Permite escolher
/// o arquivo de saida, prefixos/layers das entidades e a superficie usada
/// para calcular a declividade do talvegue.
/// </summary>
public partial class ExportarCsvBaciasSubbaciasWindow : Window
{
    private const string m_layerMtextsPadrao = "A - TEXT BACIA";
    private const string m_layerTalveguesPadrao = "HDR - TALVEGUES SUBBACIAS";
    private const string m_prefixoLayersHatchesPadrao = "HDR-BACIA";
    private const string m_prefixoRemoverBaciaPadrao = "HDR-";
    private const string m_prefixoTextoSubbaciaPadrao = "SB";
    private const double m_raioJusantePadrao = 10.0;

    /// <summary>Disparado quando o usuario clica em Exportar e a validacao passa.</summary>
    public event EventHandler<ExportarCsvBaciasSubbaciasRequest>? ConfirmarClicado;

    public ExportarCsvBaciasSubbaciasWindow(
        IReadOnlyList<string> layers,
        IReadOnlyList<string> superficies)
    {
        InitializeComponent();

        string? layerMtextsInicial = layers.FirstOrDefault(
            l => string.Equals(l, m_layerMtextsPadrao, StringComparison.OrdinalIgnoreCase));
        string? layerTalveguesInicial = layers.FirstOrDefault(
            l => string.Equals(l, m_layerTalveguesPadrao, StringComparison.OrdinalIgnoreCase));

        ConfigurarComboBoxFiltravel(LayerMTextsComboBox, layers, layerMtextsInicial ?? m_layerMtextsPadrao);
        ConfigurarComboBoxFiltravel(LayerTalveguesComboBox, layers, layerTalveguesInicial ?? m_layerTalveguesPadrao);
        ConfigurarComboBoxFiltravel(SuperficieComboBox, superficies, superficies.FirstOrDefault());

        PrefixoLayersHatchesTextBox.Text = m_prefixoLayersHatchesPadrao;
        PrefixoRemoverBaciaTextBox.Text = m_prefixoRemoverBaciaPadrao;
        PrefixoTextoSubbaciaTextBox.Text = m_prefixoTextoSubbaciaPadrao;
        RaioJusanteTextBox.Text = m_raioJusantePadrao.ToString(CultureInfo.InvariantCulture);
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
            // Nao fatal
        }
    }

    public string CaminhoCsv => CaminhoCsvTextBox.Text?.Trim() ?? string.Empty;
    public string NomeLayerMTextsSubbaciasId => LerValor(LayerMTextsComboBox);
    public string PrefixoTextoSubbacia => PrefixoTextoSubbaciaTextBox.Text?.Trim() ?? string.Empty;
    public string NomeLayerTalvegues => LerValor(LayerTalveguesComboBox);
    public string PrefixoLayersHatches => PrefixoLayersHatchesTextBox.Text?.Trim() ?? string.Empty;
    public string PrefixoRemoverDoNomeBacia => PrefixoRemoverBaciaTextBox.Text?.Trim() ?? string.Empty;
    public string NomeSuperficie => LerValor(SuperficieComboBox);

    private void AoClicarSelecionarCsv(object sender, RoutedEventArgs e)
    {
        var dialogo = new Microsoft.Win32.SaveFileDialog
        {
            Title = "Selecione o arquivo CSV de saida",
            Filter = "Arquivos CSV (*.csv)|*.csv|Todos os arquivos (*.*)|*.*",
            DefaultExt = ".csv",
            AddExtension = true,
            OverwritePrompt = true,
        };

        if (!string.IsNullOrWhiteSpace(CaminhoCsvTextBox.Text))
        {
            try
            {
                string? diretorio = Path.GetDirectoryName(CaminhoCsvTextBox.Text);
                if (!string.IsNullOrWhiteSpace(diretorio) && Directory.Exists(diretorio))
                    dialogo.InitialDirectory = diretorio;
                dialogo.FileName = Path.GetFileName(CaminhoCsvTextBox.Text);
            }
            catch
            {
                // Caminho invalido — deixa o dialogo abrir no local padrao.
            }
        }
        else
        {
            dialogo.FileName = "BACIAS_CIVIL3D.csv";
        }

        bool? resultado = dialogo.ShowDialog(this);
        if (resultado == true)
            CaminhoCsvTextBox.Text = dialogo.FileName;
    }

    private void AoClicarCancelar(object sender, RoutedEventArgs e) => Close();

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(CaminhoCsv))
        {
            MostrarAviso("Selecione o arquivo CSV de saida antes de exportar.");
            return;
        }

        if (string.IsNullOrWhiteSpace(PrefixoLayersHatches))
        {
            MostrarAviso("Informe o prefixo das layers das hatches (ex.: 'HDR-BACIA').");
            return;
        }

        if (string.IsNullOrWhiteSpace(NomeLayerMTextsSubbaciasId))
        {
            MostrarAviso("Selecione a layer dos MTexts dos IDs das subbacias.");
            return;
        }

        double raio = m_raioJusantePadrao;
        string textoRaio = RaioJusanteTextBox.Text?.Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(textoRaio))
        {
            if (!double.TryParse(textoRaio, NumberStyles.Float, CultureInfo.InvariantCulture, out raio)
                && !double.TryParse(textoRaio, NumberStyles.Float, CultureInfo.CurrentCulture, out raio))
            {
                MostrarAviso("Raio de busca invalido. Use um numero decimal com ponto ou virgula.");
                return;
            }
            if (raio <= 0)
            {
                MostrarAviso("Raio de busca deve ser maior que zero.");
                return;
            }
        }

        var request = new ExportarCsvBaciasSubbaciasRequest
        {
            CaminhoCsv = CaminhoCsv,
            NomeLayerMTextsSubbaciasId = NomeLayerMTextsSubbaciasId,
            PrefixoTextoSubbacia = PrefixoTextoSubbacia,
            NomeLayerTalvegues = NomeLayerTalvegues,
            PrefixoLayersHatches = PrefixoLayersHatches,
            PrefixoRemoverDoNomeBacia = PrefixoRemoverDoNomeBacia,
            NomeSuperficie = NomeSuperficie,
            RaioBuscaJusante = raio,
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

    private void MostrarAviso(string mensagem)
    {
        MessageBox.Show(this, mensagem, "ZagoCivil3D",
            MessageBoxButton.OK, MessageBoxImage.Warning);
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

    /// <summary>
    /// Popula a aba "Resultados" com a tabela exportada e a lista de avisos.
    /// Ativa automaticamente a aba quando ha linhas a exibir, para que o
    /// usuario veja imediatamente o resultado sem ter que clicar na aba.
    /// </summary>
    public void PopularResultados(ExportarCsvBaciasSubbaciasResultado resultado)
    {
        // Resumo no topo da aba
        if (resultado.Linhas.Count == 0)
        {
            ResumoResultadosTextBlock.Text = resultado.MensagensErro.Count > 0
                ? "Nenhuma linha gerada. Verifique os erros abaixo ou na linha de comando."
                : "Nenhuma linha gerada.";
        }
        else
        {
            ResumoResultadosTextBlock.Text =
                $"{resultado.Linhas.Count} linha(s) exportada(s) — " +
                $"{resultado.TotalLinhasComAviso} com avisos (REVISAR).";
        }

        if (!string.IsNullOrWhiteSpace(resultado.CaminhoCsvGerado))
        {
            ArquivoGeradoTextBlock.Text = $"Arquivo: {resultado.CaminhoCsvGerado}";
            ArquivoGeradoTextBlock.Visibility = Visibility.Visible;
        }
        else
        {
            ArquivoGeradoTextBlock.Visibility = Visibility.Collapsed;
        }

        // Tabela
        LinhasDataGrid.ItemsSource = resultado.Linhas;

        // Avisos: uma entrada por linha com qualquer VERF != OK, incluindo
        // tambem os erros globais (MensagensErro) como primeiro grupo.
        var grupos = new List<GrupoAviso>();

        if (resultado.MensagensErro.Count > 0)
        {
            grupos.Add(new GrupoAviso(
                "Erros gerais",
                resultado.MensagensErro.ToList()));
        }

        foreach (LinhaCsvBaciasSubbacias linha in resultado.Linhas.Where(l => l.TemAviso))
        {
            string titulo = MontarTituloAviso(linha);
            var mensagens = new List<string>();
            if (linha.VerfSb != LinhaCsvBaciasSubbacias.VerfOk)
                mensagens.Add($"VERF_SB: {linha.VerfSb}");
            if (linha.VerfA != LinhaCsvBaciasSubbacias.VerfOk)
                mensagens.Add($"VERF_A: {linha.VerfA}");
            if (linha.VerfL != LinhaCsvBaciasSubbacias.VerfOk)
                mensagens.Add($"VERF_L: {linha.VerfL}");
            if (linha.VerfS != LinhaCsvBaciasSubbacias.VerfOk)
                mensagens.Add($"VERF_S: {linha.VerfS}");
            grupos.Add(new GrupoAviso(titulo, mensagens));
        }

        if (grupos.Count > 0)
        {
            AvisosItemsControl.ItemsSource = grupos;
            AvisosCabecalhoTextBlock.Text = resultado.MensagensErro.Count > 0
                ? $"Avisos ({resultado.TotalLinhasComAviso}) e erros ({resultado.MensagensErro.Count})"
                : $"Avisos (REVISAR) — {resultado.TotalLinhasComAviso} subbacia(s)";
            AvisosBorder.Visibility = Visibility.Visible;
        }
        else
        {
            AvisosItemsControl.ItemsSource = null;
            AvisosBorder.Visibility = Visibility.Collapsed;
        }

        // Ativa aba de resultados quando ha algo a mostrar.
        if (resultado.Linhas.Count > 0 || grupos.Count > 0)
            AbasPrincipais.SelectedItem = AbaResultadosItem;
    }

    private static string MontarTituloAviso(LinhaCsvBaciasSubbacias linha)
    {
        string subbacia = string.IsNullOrWhiteSpace(linha.Subbacia) ? "(sem ID)" : linha.Subbacia;
        string bacia = string.IsNullOrWhiteSpace(linha.Bacia) ? "?" : linha.Bacia;
        return $"Subbacia {subbacia} (bacia {bacia})";
    }

    /// <summary>Par titulo/mensagens para renderizar cada grupo de avisos.</summary>
    private sealed class GrupoAviso
    {
        public GrupoAviso(string titulo, List<string> mensagens)
        {
            Titulo = titulo;
            Mensagens = mensagens;
        }

        public string Titulo { get; }
        public List<string> Mensagens { get; }
    }

    private static string LerValor(ComboBox combo)
    {
        if (combo.SelectedItem is string selecionado && !string.IsNullOrWhiteSpace(selecionado))
            return selecionado.Trim();
        return combo.Text?.Trim() ?? string.Empty;
    }

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
                if (todos.Contains(inicial, StringComparer.OrdinalIgnoreCase))
                    combo.SelectedItem = todos.First(
                        i => string.Equals(i, inicial, StringComparison.OrdinalIgnoreCase));
                else
                    combo.Text = inicial;
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
