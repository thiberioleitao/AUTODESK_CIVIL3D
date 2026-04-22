using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using ZagoCivil3D.Models;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ZagoCivil3D.Views;

/// <summary>
/// Janela modeless para criar corredores vazios a partir de todos os
/// alinhamentos do desenho. Permanece aberta apos execucao para facilitar
/// repeticoes com prefixos/sufixos diferentes.
/// </summary>
public partial class CriarCorredoresAlinhamentosWindow : Window
{
    /// <summary>
    /// Evento disparado quando o usuario clica em Executar e a validacao passa.
    /// </summary>
    public event EventHandler<CriarCorredoresAlinhamentosRequest>? ConfirmarClicado;

    private readonly string? m_nomePrimeiroAlinhamento;

    public CriarCorredoresAlinhamentosWindow(IReadOnlyList<string> nomesAlinhamentos)
    {
        InitializeComponent();

        m_nomePrimeiroAlinhamento = nomesAlinhamentos?.FirstOrDefault();

        PrefixoTextBox.TextChanged += (s, e) => AtualizarExemplo();
        SufixoTextBox.TextChanged += (s, e) => AtualizarExemplo();

        AtualizarExemplo();
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

    public string PrefixoNome => PrefixoTextBox.Text?.Trim() ?? string.Empty;

    public string SufixoNome => SufixoTextBox.Text?.Trim() ?? string.Empty;

    public bool IgnorarExistentes => IgnorarExistentesCheckBox.IsChecked == true;

    private void AoClicarCancelar(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void AoClicarConfirmar(object sender, RoutedEventArgs e)
    {
        var request = new CriarCorredoresAlinhamentosRequest
        {
            PrefixoNome = PrefixoNome,
            SufixoNome = SufixoNome,
            IgnorarExistentes = IgnorarExistentes
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
    /// Atualiza o preview do nome resultante aplicando prefixo/sufixo ao nome
    /// do primeiro alinhamento encontrado, para dar feedback visual imediato.
    /// </summary>
    private void AtualizarExemplo()
    {
        if (ExemploTextBlock == null) return;

        if (string.IsNullOrWhiteSpace(m_nomePrimeiroAlinhamento))
        {
            ExemploTextBlock.Text = "(nenhum alinhamento no desenho)";
            return;
        }

        ExemploTextBlock.Text = PrefixoNome + m_nomePrimeiroAlinhamento + SufixoNome;
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
}
