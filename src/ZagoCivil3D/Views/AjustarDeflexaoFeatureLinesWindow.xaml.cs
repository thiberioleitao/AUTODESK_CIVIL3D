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
/// Janela modeless para o ajuste iterativo de deflexao de feature lines
/// (porta da rotina Dynamo TER). Permanece aberta apos a execucao para
/// permitir rodar com parametros diferentes sem reabrir o comando.
/// </summary>
public partial class AjustarDeflexaoFeatureLinesWindow : Window
{
    private static readonly CultureInfo s_cultura = CultureInfo.InvariantCulture;

    /// <summary>
    /// Evento disparado quando o usuario clica em Executar e a validacao passa.
    /// </summary>
    public event EventHandler<AjustarDeflexaoFeatureLinesRequest>? ConfirmarClicado;

    public AjustarDeflexaoFeatureLinesWindow(
        IReadOnlyList<string> sites,
        IReadOnlyList<string> layersDesenho)
    {
        InitializeComponent();

        // Inclui uma opcao vazia para permitir "todos os sites".
        var listaSites = new List<string> { string.Empty };
        listaSites.AddRange(sites);
        ConfigurarComboBoxFiltravel(SiteComboBox, listaSites, sites.FirstOrDefault() ?? string.Empty);

        // Layers do marcador (uma por status). Cada ComboBox e alimentado com as layers
        // existentes no desenho, mais o nome padrao correspondente ao status (caso ele
        // nao exista ainda, sera criado no momento do uso).
        ConfigurarComboBoxLayer(
            LayerAjustadoComboBox,
            layersDesenho,
            Services.MarcadoresVisuaisDeflexaoService.LayerPadraoAjustado);
        ConfigurarComboBoxLayer(
            LayerFalhaComboBox,
            layersDesenho,
            Services.MarcadoresVisuaisDeflexaoService.LayerPadraoFalha);
        ConfigurarComboBoxLayer(
            LayerInalteradoComboBox,
            layersDesenho,
            Services.MarcadoresVisuaisDeflexaoService.LayerPadraoInalterado);

        // Valores padrao.
        DeflexaoLimiteTextBox.Text = "0.0129";
        PassoTextBox.Text = "0.01";
        PassoMaximoTextBox.Text = "0.50";
        IndiceMaximoTextBox.Text = "1000";
        PassadasTextBox.Text = "2";
        IteracoesPorPontoTextBox.Text = "200";
        TamanhoMarcadorTextBox.Text = "1.0";
        FiltroNomeTextBox.Text = string.Empty;
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

    public string NomeSite => LerValor(SiteComboBox);

    public string FiltroNome => FiltroNomeTextBox.Text?.Trim() ?? string.Empty;

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        if (!TentarLerParametros(out AjustarDeflexaoFeatureLinesRequest? request) || request == null)
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
    /// Popula as tabelas da aba "Resultados" a partir do resultado da execucao
    /// e alterna automaticamente para essa aba. Chamado pelo command apos o
    /// service terminar de processar.
    /// </summary>
    public void AtualizarRelatorio(AjustarDeflexaoFeatureLinesResultado resultado)
    {
        // Guarda uma copia completa para suportar o filtro de "apenas violacoes".
        m_todosOsPontos = resultado.RelatorioPontos.ToList();

        GridResumoFL.ItemsSource = null;
        GridResumoFL.ItemsSource = resultado.RelatorioFeatureLines;

        AplicarFiltroDetalhes();

        ResumoAgregadoTextBlock.Text = MontarTituloAgregado(resultado);
        ResumoAgregadoDetalheTextBlock.Text = MontarDetalheAgregado(resultado);

        TabResultados.IsEnabled = true;

        // Alterna automaticamente para a aba de resultados quando ha o que mostrar.
        if (resultado.RelatorioFeatureLines.Count > 0)
            TabResultados.IsSelected = true;
    }

    private List<RelatorioPontoPI> m_todosOsPontos = new();

    private void AoMudarFiltroDetalhes(object sender, RoutedEventArgs e)
    {
        AplicarFiltroDetalhes();
    }

    private void AplicarFiltroDetalhes()
    {
        if (GridDetalhesPI == null)
            return;

        bool apenasViolacoes = FiltrarSoViolacoesCheckBox?.IsChecked == true;

        IEnumerable<RelatorioPontoPI> fonte = m_todosOsPontos;
        if (apenasViolacoes)
            fonte = fonte.Where(p => p.Status == StatusAjustePI.ViolacaoRemanescente
                                     || p.Status == StatusAjustePI.FalhaGeometrica);

        GridDetalhesPI.ItemsSource = null;
        GridDetalhesPI.ItemsSource = fonte.ToList();
    }

    private static string MontarTituloAgregado(AjustarDeflexaoFeatureLinesResultado resultado)
    {
        if (resultado.TotalFeatureLinesConsideradas == 0)
            return "Nenhuma feature line processada.";

        if (resultado.TotalViolacoesIniciais == 0)
            return $"{resultado.TotalFeatureLinesConsideradas} feature line(s) sem violacoes de deflexao.";

        if (resultado.TotalViolacoesFinaisInternas == 0 && resultado.TotalViolacoesFinaisFronteira == 0)
            return $"Todas as {resultado.TotalViolacoesIniciais} violacoes foram corrigidas.";

        return $"Violacoes: {resultado.TotalViolacoesIniciais} inicial(is) -> {resultado.TotalViolacoesFinais} restante(s)" +
               $" (internas {resultado.TotalViolacoesFinaisInternas}, fronteira {resultado.TotalViolacoesFinaisFronteira}).";
    }

    private static string MontarDetalheAgregado(AjustarDeflexaoFeatureLinesResultado resultado)
    {
        string deltaMedio = resultado.DeltaZMedio.ToString("F4", s_cultura);
        string deltaMax = resultado.DeltaZMaximo.ToString("F4", s_cultura);
        return $"{resultado.TotalEstacoesAlteradas} estacao(oes) distintas alterada(s) em " +
               $"{resultado.TotalFeatureLinesAjustadas} feature line(s), com {resultado.TotalIteracoesAplicadas} " +
               $"iteracao(oes) aplicada(s). |dZ| medio: {deltaMedio} m, maximo: {deltaMax} m. " +
               $"Violacoes de fronteira (PI 1 / PI n-2) dependem de feature lines vizinhas e podem " +
               $"nao ser corrigiveis localmente.";
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
    /// Le e valida os parametros numericos informados no formulario.
    /// </summary>
    private bool TentarLerParametros(out AjustarDeflexaoFeatureLinesRequest? request)
    {
        request = null;

        if (!TentarLerDouble(DeflexaoLimiteTextBox.Text, "Deflexao limite", permitirNegativo: false, out double deflexaoLimite))
            return false;

        if (deflexaoLimite <= 0)
        {
            AvisarValor("Deflexao limite", "deve ser maior que zero.");
            return false;
        }

        if (!TentarLerDouble(PassoTextBox.Text, "Passo minimo / precisao", permitirNegativo: false, out double passoMin))
            return false;

        if (passoMin <= 0)
        {
            AvisarValor("Passo minimo / precisao", "deve ser maior que zero.");
            return false;
        }

        if (!TentarLerDouble(PassoMaximoTextBox.Text, "Passo maximo dinamico", permitirNegativo: false, out double passoMax))
            return false;

        if (passoMax < passoMin)
        {
            AvisarValor("Passo maximo dinamico", $"deve ser maior ou igual ao passo minimo ({passoMin}).");
            return false;
        }

        if (!TentarLerInt(IndiceMaximoTextBox.Text, "Indice maximo de PI", out int indiceMaximo))
            return false;

        if (indiceMaximo < 1)
        {
            AvisarValor("Indice maximo de PI", "deve ser pelo menos 1.");
            return false;
        }

        if (!TentarLerInt(PassadasTextBox.Text, "Numero de passadas", out int passadas))
            return false;

        if (passadas < 1 || passadas > 10)
        {
            AvisarValor("Numero de passadas", "deve estar entre 1 e 10.");
            return false;
        }

        if (!TentarLerInt(IteracoesPorPontoTextBox.Text, "Limite de iteracoes por ponto", out int iteracoesPorPonto))
            return false;

        if (iteracoesPorPonto < 1)
        {
            AvisarValor("Limite de iteracoes por ponto", "deve ser pelo menos 1.");
            return false;
        }

        bool criarMarcadores = CriarMarcadoresCheckBox.IsChecked == true;
        double alturaTexto = 1.0;
        string nomeLayerAjustado = string.Empty;
        string nomeLayerFalha = string.Empty;
        string nomeLayerInalterado = string.Empty;
        if (criarMarcadores)
        {
            if (!TentarLerDouble(TamanhoMarcadorTextBox.Text, "Altura do texto", permitirNegativo: false, out alturaTexto))
                return false;

            if (alturaTexto <= 0)
            {
                AvisarValor("Altura do texto", "deve ser maior que zero.");
                return false;
            }

            nomeLayerAjustado = LerValor(LayerAjustadoComboBox);
            nomeLayerFalha = LerValor(LayerFalhaComboBox);
            nomeLayerInalterado = LerValor(LayerInalteradoComboBox);
        }

        ModoOtimizacaoDeflexao modo = LerModoOtimizacao();

        request = new AjustarDeflexaoFeatureLinesRequest
        {
            NomeSite = NomeSite,
            FiltroNome = FiltroNome,
            DeflexaoLimite = deflexaoLimite,
            PassoIncrementalAjusteCota = passoMin,
            PassoMaximoDinamico = passoMax,
            IndiceMaximo = indiceMaximo,
            QuantidadePassadas = passadas,
            NumeroMaximoIteracoesPorPonto = iteracoesPorPonto,
            ModoOtimizacao = modo,
            CriarMarcadoresNoDesenho = criarMarcadores,
            NomeLayerMarcadorAjustado = nomeLayerAjustado,
            NomeLayerMarcadorFalha = nomeLayerFalha,
            NomeLayerMarcadorInalterado = nomeLayerInalterado,
            AlturaTextoMarcador = alturaTexto,
        };

        return true;
    }

    /// <summary>
    /// Alimenta um ComboBox de layer com as layers do desenho, garantindo que o nome padrao
    /// tambem esteja disponivel (mesmo que ainda nao exista no desenho — sera criado no uso).
    /// </summary>
    private static void ConfigurarComboBoxLayer(
        System.Windows.Controls.ComboBox combo,
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
    /// Le o modo de otimizacao selecionado no ComboBox. A identificacao e feita pela Tag
    /// do ComboBoxItem para independer do texto apresentado ao usuario.
    /// </summary>
    private ModoOtimizacaoDeflexao LerModoOtimizacao()
    {
        if (ModoComboBox.SelectedItem is System.Windows.Controls.ComboBoxItem item
            && item.Tag is string tag
            && string.Equals(tag, "PontoCentral", StringComparison.OrdinalIgnoreCase))
        {
            return ModoOtimizacaoDeflexao.PontoCentral;
        }

        return ModoOtimizacaoDeflexao.MultiPonto;
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

    private bool TentarLerInt(string? texto, string nomeCampo, out int valor)
    {
        valor = 0;
        string normalizado = (texto ?? string.Empty).Trim();

        if (!int.TryParse(normalizado, NumberStyles.Integer, s_cultura, out valor))
        {
            AvisarValor(nomeCampo, "deve ser um numero inteiro.");
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
    /// Configura um ComboBox editavel com filtro em tempo real por substring.
    /// Swap direto do ItemsSource preservando texto e cursor (ver pattern_searchable_combobox).
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
