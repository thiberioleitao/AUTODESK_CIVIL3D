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
/// Janela modeless para o comando "Criar Catchments a partir das Hatches".
/// Replica a rotina Dynamo de criacao de subbacias a partir de hatches +
/// MTexts + polilinhas de talvegue.
///
/// Permanece aberta apos a execucao para permitir rodar com parametros
/// diferentes sem reabrir o comando.
/// </summary>
public partial class CriarCatchmentsDeHatchsWindow : Window
{
    /// <summary>
    /// Evento disparado quando o usuario clica em Executar e a validacao passa.
    /// </summary>
    public event EventHandler<CriarCatchmentsDeHatchsRequest>? ConfirmarClicado;

    public CriarCatchmentsDeHatchsWindow(
        IReadOnlyList<string> layersDesenho,
        IReadOnlyList<string> superficies,
        IReadOnlyList<string> gruposCatchment,
        IReadOnlyList<string> estilosCatchment)
    {
        InitializeComponent();

        // Permite tanto escolher uma layer existente quanto digitar uma nova.
        // Valores padrao alinhados com a rotina Dynamo original do projeto Zago.
        ConfigurarComboBoxFiltravel(
            LayerHatchesComboBox,
            layersDesenho,
            EscolherInicialComPadrao(layersDesenho, "HDR-BACIA"));
        ConfigurarComboBoxFiltravel(
            LayerMTextsComboBox,
            layersDesenho,
            EscolherInicialComPadrao(layersDesenho, "HDR-ANOT IDS SUBBACIAS"));

        // Talvegue permite entrada vazia (sem flow path).
        var layersComVazio = new List<string> { string.Empty };
        layersComVazio.AddRange(layersDesenho);
        ConfigurarComboBoxFiltravel(
            LayerTalveguesComboBox,
            layersComVazio,
            EscolherInicialComPadrao(layersDesenho, "HDR-TALVEGUES SUBBACIAS"));

        // Grupo permite digitar um novo nome (sera criado se nao existir).
        ConfigurarComboBoxFiltravel(
            GrupoCatchmentComboBox,
            gruposCatchment,
            gruposCatchment.FirstOrDefault() ?? "HDR-BACIA");

        // Estilo permite entrada vazia (usa o primeiro disponivel).
        var estilosComVazio = new List<string> { string.Empty };
        estilosComVazio.AddRange(estilosCatchment);
        ConfigurarComboBoxFiltravel(
            EstiloCatchmentComboBox,
            estilosComVazio,
            string.Empty);

        // Superficie permite entrada vazia (sem superficie de referencia).
        var superficiesComVazio = new List<string> { string.Empty };
        superficiesComVazio.AddRange(superficies);
        ConfigurarComboBoxFiltravel(
            SuperficieComboBox,
            superficiesComVazio,
            superficies.FirstOrDefault() ?? string.Empty);

        PrefixoBaciaTextBox.Text = "SB";
        SeparadorNomeTextBox.Text = "-";
    }

    /// <summary>
    /// Define a janela do AutoCAD como owner para garantir foco de teclado nos campos editaveis.
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

    public string NomeLayerHatches => LerValor(LayerHatchesComboBox);
    public string NomeLayerMTextsIds => LerValor(LayerMTextsComboBox);
    public string NomeLayerTalvegues => LerValor(LayerTalveguesComboBox);
    public string NomeGrupoCatchment => LerValor(GrupoCatchmentComboBox);
    public string NomeEstiloCatchment => LerValor(EstiloCatchmentComboBox);
    public string NomeSuperficie => LerValor(SuperficieComboBox);
    public string PrefixoBacia => PrefixoBaciaTextBox.Text?.Trim() ?? string.Empty;
    public string SeparadorNome => SeparadorNomeTextBox.Text ?? "-";
    public bool ConfigurarFlowPath => ConfigurarFlowPathCheckBox.IsChecked == true;
    public bool SubstituirExistentes => SubstituirExistentesCheckBox.IsChecked == true;

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        // Validacao de borda: campos obrigatorios.
        if (string.IsNullOrWhiteSpace(NomeLayerHatches))
        {
            AvisarUsuario("Selecione ou digite o nome da layer das hatches.");
            return;
        }

        if (string.IsNullOrWhiteSpace(NomeLayerMTextsIds))
        {
            AvisarUsuario("Selecione ou digite o nome da layer dos MTexts de ID.");
            return;
        }

        if (string.IsNullOrWhiteSpace(NomeGrupoCatchment))
        {
            AvisarUsuario("Informe o nome do Catchment Group.");
            return;
        }

        var request = new CriarCatchmentsDeHatchsRequest
        {
            NomeLayerHatches = NomeLayerHatches,
            NomeLayerMTextsIds = NomeLayerMTextsIds,
            NomeLayerTalvegues = NomeLayerTalvegues,
            NomeGrupoCatchment = NomeGrupoCatchment,
            PrefixoBacia = PrefixoBacia,
            SeparadorNome = SeparadorNome,
            NomeSuperficie = NomeSuperficie,
            NomeEstiloCatchment = NomeEstiloCatchment,
            ConfigurarFlowPath = ConfigurarFlowPath,
            SubstituirExistentes = SubstituirExistentes,
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

    private void AvisarUsuario(string mensagem)
    {
        MessageBox.Show(this, mensagem, "ZagoCivil3D", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    /// <summary>
    /// Retorna <paramref name="padrao"/> se estiver presente na lista (case-insensitive),
    /// caso contrario retorna o primeiro item da lista (ou o proprio padrao, como fallback).
    /// </summary>
    private static string EscolherInicialComPadrao(IReadOnlyList<string> itens, string padrao)
    {
        if (itens.Any(i => string.Equals(i, padrao, StringComparison.OrdinalIgnoreCase)))
            return padrao;
        return itens.FirstOrDefault() ?? padrao;
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
                // Seleciona pelo match exato se possivel; caso contrario deixa o texto como esta.
                string? match = todos.FirstOrDefault(i => string.Equals(i, inicial, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                    combo.SelectedItem = match;
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
