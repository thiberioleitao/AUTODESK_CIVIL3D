// RibbonInitializer.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using System;
using System.Linq;
using System.Runtime.Versioning;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ZagoCivil3D.Ribbon
{
    [SupportedOSPlatform("windows")]
    /// <summary>
    /// Centraliza a criação da aba personalizada no Ribbon do Civil 3D.
    /// Mantém lógica idempotente para evitar abas duplicadas em recargas.
    /// </summary>
    public static class RibbonInitializer
    {
        private const string m_tabDrenagemId = "ZAGO_DRENAGEM_TAB";
        private const string m_panelCriarAlinhamentosId = "ZAGO_DRENAGEM_CRIAR_ALINHAMENTOS_PANEL";
        private const string m_panelCriarPerfisId = "ZAGO_DRENAGEM_CRIAR_PERFIS_PANEL";
        private const string m_panelCriarCorredoresId = "ZAGO_DRENAGEM_CRIAR_CORREDORES_PANEL";
        private const string m_panelCriarRegioesId = "ZAGO_DRENAGEM_CRIAR_REGIOES_PANEL";
        private const string m_panelCriarCaixasId = "ZAGO_DRENAGEM_CRIAR_CAIXAS_PANEL";
        private const string m_panelExportarId = "ZAGO_DRENAGEM_EXPORTAR_PANEL";
        private const string m_panelModificarId = "ZAGO_DRENAGEM_MODIFICAR_PANEL";
        private const string m_panelAnotarId = "ZAGO_DRENAGEM_ANOTAR_PANEL";
        private const string m_panelDeletarId = "ZAGO_DRENAGEM_DELETAR_PANEL";
        private const string m_tabTerraplenagemId = "ZAGO_TERRAPLENAGEM_TAB";
        private const string m_panelDummyId = "ZAGO_TERRAPLENAGEM_DUMMY_PANEL";
        private const string m_prefixoComandoDummy = "DUMMY_PRINT:";
        private static bool m_aguardandoRibbon;

        public static void InicializarRibbon()
        {
            try
            {
                var documento = Application.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage("\n[ZagoCivil3D] InitializeRibbon chamado.");
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Ribbon null? {ComponentManager.Ribbon == null}");

                if (ComponentManager.Ribbon != null)
                {
                    documento?.Editor.WriteMessage("\n[ZagoCivil3D] Ribbon já disponível.");
                    CriarRibbon(ComponentManager.Ribbon);
                    return;
                }

                // O Ribbon pode ainda não existir durante o NETLOAD.
                // Nesse caso aguardamos o evento de inicialização da UI.
                if (m_aguardandoRibbon)
                {
                    documento?.Editor.WriteMessage("\n[ZagoCivil3D] Já aguardando ItemInitialized. Nenhuma nova inscrição realizada.");
                    return;
                }

                documento?.Editor.WriteMessage("\n[ZagoCivil3D] Ribbon ainda não disponível. Inscrevendo ItemInitialized.");
                ComponentManager.ItemInitialized += AoInicializarItem;
                m_aguardandoRibbon = true;
            }
            catch (Exception ex)
            {
                var documento = Application.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Erro em InitializeRibbon: {ex.Message}");
            }
        }

        private static void AoInicializarItem(object? remetente, RibbonItemEventArgs argumentos)
        {
            try
            {
                var documento = Application.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] OnItemInitialized disparado. Item: {argumentos.Item?.Id ?? "(sem id)"}");

                if (ComponentManager.Ribbon == null)
                {
                    documento?.Editor.WriteMessage("\n[ZagoCivil3D] OnItemInitialized: Ribbon ainda nula.");
                    return;
                }

                CriarRibbon(ComponentManager.Ribbon);
                ComponentManager.ItemInitialized -= AoInicializarItem;
                m_aguardandoRibbon = false;
                documento?.Editor.WriteMessage("\n[ZagoCivil3D] ItemInitialized removido.");
            }
            catch (Exception ex)
            {
                var documento = Application.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Erro em OnItemInitialized: {ex.Message}");
            }
        }

        private static void CriarRibbon(RibbonControl ribbon)
        {
            try
            {
                var documento = Application.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage("\n[ZagoCivil3D] CreateRibbon chamado.");
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Total de abas atuais: {ribbon.Tabs.Count}");

                RibbonTab abaDrenagem = ObterOuCriarAba(ribbon, m_tabDrenagemId, "ZAGO - DRENAGEM");
                RibbonPanelSource painelCriarAlinhamentos = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarAlinhamentosId, "CRIAR - ALINHAMENTOS");
                AdicionarBotaoComando(painelCriarAlinhamentos, "ZAGO_CRIAR_ALINH_POR_POLI", "Alinhamentos\npor Polilinha", "ZAGO_CRIAR_ALINHAMENTOS_POR_POLILINHA ", "AL");
                AdicionarBotaoGrande(painelCriarAlinhamentos, "ZAGO_CRIAR_PONTOS_CRUZ", "Pontos nos\nCruzamentos", "CRIAR - ALINHAMENTOS > PONTOS NOS CRUZAMENTOS ENTRE ALINHAMENTOS", "PX");

                RibbonPanelSource painelCriarPerfis = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarPerfisId, "CRIAR - PERFIS");
                AdicionarBotaoGrande(painelCriarPerfis, "ZAGO_CRIAR_PERFIL_PROJ", "Perfis de\nProjeto", "CRIAR - PERFIS > PERFIS DE PROJETO A PARTIR DE ALINHAMENTOS", "PP");
                AdicionarBotaoGrande(painelCriarPerfis, "ZAGO_CRIAR_PERFIL_TN", "Perfis TN e\nTerraplenagem", "CRIAR - PERFIS > PERFIS DO TERRENO NATURAL E TERRAPLENAGEM", "TN");
                AdicionarBotaoGrande(painelCriarPerfis, "ZAGO_CRIAR_PROFILE_VIEW", "Profile\nView", "CRIAR - PERFIS > PROFILE VIEW", "PV");

                RibbonPanelSource painelCriarCorredores = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarCorredoresId, "CRIAR - CORREDORES");
                AdicionarBotaoGrande(painelCriarCorredores, "ZAGO_CRIAR_CORREDORES", "Criar\nCorredores", "CRIAR - CORREDORES > FUNCOES DE CRIACAO DE CORREDORES", "CO");

                RibbonPanelSource painelCriarRegioes = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarRegioesId, "CRIAR - REGIOES");
                AdicionarBotaoGrande(painelCriarRegioes, "ZAGO_CRIAR_REGIOES", "Criar\nRegioes", "CRIAR - REGIOES > FUNCOES DE CRIACAO DE REGIOES", "RG");

                RibbonPanelSource painelCriarCaixas = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarCaixasId, "CRIAR - CAIXAS");
                AdicionarBotaoGrande(painelCriarCaixas, "ZAGO_CRIAR_CAIXAS", "Criar\nCaixas", "CRIAR - CAIXAS > FUNCOES DE CRIACAO DE CAIXAS", "CX");

                RibbonPanelSource painelExportar = ObterOuCriarPainelFonte(abaDrenagem, m_panelExportarId, "EXPORTAR");
                AdicionarBotaoGrande(painelExportar, "ZAGO_EXPORTAR_CSV_BACIAS", "CSV Bacias e\nSubbacias", "EXPORTAR > CSV BACIAS/SUBBACIAS (ID, AREA, TALVEGUE, DECLIVIDADE, ID_JUSANTE)", "SB");
                AdicionarBotaoGrande(painelExportar, "ZAGO_EXPORTAR_CSV_CANAIS", "CSV Canais\ne Bueiros", "EXPORTAR > CSV CANAIS E BUEIROS", "CE");

                RibbonPanelSource painelModificar = ObterOuCriarPainelFonte(abaDrenagem, m_panelModificarId, "MODIFICAR");
                AdicionarBotaoGrande(painelModificar, "ZAGO_MODIFICAR_DUMMY", "Modificar\n(Dummy)", "MODIFICAR > FUNCOES EM DEFINICAO", "MD");

                RibbonPanelSource painelAnotar = ObterOuCriarPainelFonte(abaDrenagem, m_panelAnotarId, "ANOTAR");
                AdicionarBotaoGrande(painelAnotar, "ZAGO_ANOTAR_LABELS_CORR", "Labels\nCorredores", "ANOTAR > ADICIONAR LABELS DOS TRECHOS/REGIOES DOS CORREDORES EM PLANTA", "LB");

                RibbonPanelSource painelDeletar = ObterOuCriarPainelFonte(abaDrenagem, m_panelDeletarId, "DELETAR");
                AdicionarBotaoGrande(painelDeletar, "ZAGO_DELETAR_DUMMY", "Deletar\n(Dummy)", "DELETAR > FUNCOES EM DEFINICAO", "DL");

                RibbonTab abaTerraplenagem = ObterOuCriarAba(ribbon, m_tabTerraplenagemId, "ZAGO - TERRAPLENAGEM");
                RibbonPanelSource painelDummyTerraplenagem = ObterOuCriarPainelFonte(abaTerraplenagem, m_panelDummyId, "BOTÃO DUMMY");
                AdicionarBotaoGrande(painelDummyTerraplenagem, "ZAGO_TERRAPL_DUMMY", "Botão\nDummy", "TERRAPLENAGEM > BOTAO DUMMY", "TP");

                ribbon.ActiveTab = abaDrenagem;

                documento?.Editor.WriteMessage("\n[ZagoCivil3D] Aba e botão criados com sucesso.");
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Aba ativa: {ribbon.ActiveTab?.Id ?? "(nenhuma)"}");
            }
            catch (Exception ex)
            {
                var documento = Application.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Erro em CreateRibbon: {ex.Message}");
            }
        }

        private static RibbonTab ObterOuCriarAba(RibbonControl ribbon, string idAba, string tituloAba)
        {
            RibbonTab? abaExistente = ribbon.Tabs.FirstOrDefault(t => t.Id == idAba);
            if (abaExistente != null)
                return abaExistente;

            var aba = new RibbonTab
            {
                Title = tituloAba,
                Id = idAba
            };

            ribbon.Tabs.Add(aba);
            return aba;
        }

        private static RibbonPanelSource ObterOuCriarPainelFonte(
            RibbonTab aba,
            string idPainel,
            string tituloPainel)
        {
            RibbonPanel? painelExistente = aba.Panels.FirstOrDefault(p => p.Source?.Id == idPainel);
            if (painelExistente?.Source != null)
                return painelExistente.Source;

            var fontePainel = new RibbonPanelSource
            {
                Title = tituloPainel,
                Id = idPainel
            };

            var painel = new RibbonPanel
            {
                Source = fontePainel
            };

            aba.Panels.Add(painel);
            return fontePainel;
        }

        private static void AdicionarBotaoGrande(
            RibbonPanelSource fontePainel,
            string idBotao,
            string textoBotao,
            string mensagemDummy,
            string siglaIcone)
        {
            AdicionarBotao(fontePainel, idBotao, textoBotao, m_prefixoComandoDummy + mensagemDummy, siglaIcone);
        }

        private static void AdicionarBotaoComando(
            RibbonPanelSource fontePainel,
            string idBotao,
            string textoBotao,
            string nomeComando,
            string siglaIcone)
        {
            AdicionarBotao(fontePainel, idBotao, textoBotao, nomeComando, siglaIcone);
        }

        private static void AdicionarBotao(
            RibbonPanelSource fontePainel,
            string idBotao,
            string textoBotao,
            string parametroComando,
            string siglaIcone)
        {
            bool botaoExiste = fontePainel.Items
                .OfType<RibbonButton>()
                .Any(b => string.Equals(b.Id, idBotao, StringComparison.OrdinalIgnoreCase));

            if (botaoExiste)
                return;

            var botao = new RibbonButton
            {
                Id = idBotao,
                Text = textoBotao,
                ShowText = true,
                Size = RibbonItemSize.Large,
                Orientation = System.Windows.Controls.Orientation.Vertical,
                LargeImage = CriarIcone(siglaIcone),
                CommandParameter = parametroComando,
                CommandHandler = new ManipuladorComandoRibbon()
            };

            fontePainel.Items.Add(botao);
        }

        private static ImageSource CriarIcone(string sigla)
        {
            const int largura = 32;
            const int altura = 32;

            var visual = new DrawingVisual();
            using DrawingContext contexto = visual.RenderOpen();

            Color corFundo = ObterCorIcone(sigla);
            var retangulo = new Rect(0, 0, largura, altura);
            var pincelFundo = new SolidColorBrush(corFundo);
            pincelFundo.Freeze();

            var pincelBorda = new SolidColorBrush(Color.FromRgb(33, 37, 43));
            pincelBorda.Freeze();

            contexto.DrawRoundedRectangle(pincelFundo, new Pen(pincelBorda, 1), retangulo, 6, 6);

            double pixelsPorDip = VisualTreeHelper.GetDpi(visual).PixelsPerDip;
            var texto = new FormattedText(
                sigla,
                System.Globalization.CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight,
                new Typeface(new FontFamily("Segoe UI Semibold"), FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal),
                12,
                Brushes.White,
                pixelsPorDip);

            Point pontoTexto = new(
                (largura - texto.Width) / 2,
                (altura - texto.Height) / 2 - 1);

            contexto.DrawText(texto, pontoTexto);

            var bitmap = new RenderTargetBitmap(largura, altura, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(visual);
            bitmap.Freeze();

            return bitmap;
        }

        private static Color ObterCorIcone(string sigla)
        {
            return sigla switch
            {
                "AL" => Color.FromRgb(0, 120, 212),
                "PX" => Color.FromRgb(90, 90, 220),
                "PP" => Color.FromRgb(0, 153, 102),
                "TN" => Color.FromRgb(46, 139, 87),
                "PV" => Color.FromRgb(0, 153, 153),
                "CO" => Color.FromRgb(128, 0, 128),
                "RG" => Color.FromRgb(184, 134, 11),
                "CX" => Color.FromRgb(210, 105, 30),
                "SB" => Color.FromRgb(220, 20, 60),
                "CE" => Color.FromRgb(178, 34, 34),
                "MD" => Color.FromRgb(255, 140, 0),
                "LB" => Color.FromRgb(72, 61, 139),
                "DL" => Color.FromRgb(169, 0, 0),
                "TP" => Color.FromRgb(105, 105, 105),
                _ => Color.FromRgb(96, 96, 96)
            };
        }
    }

    [SupportedOSPlatform("windows")]
    /// <summary>
    /// Encaminha o clique do botão para um comando de texto do AutoCAD.
    /// </summary>
    public class ManipuladorComandoRibbon : System.Windows.Input.ICommand
    {
        private const string m_comandoPadrao = "ZAGO_CRIAR_ALINHAMENTOS_POR_POLILINHA ";
        private const string m_prefixoComandoDummy = "DUMMY_PRINT:";

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => true;

        public void Execute(object? parameter)
        {
            var documento = Application.DocumentManager.MdiActiveDocument;
            documento?.Editor.WriteMessage("\n[ZagoCivil3D] Botão do ribbon clicado.");

            if (documento == null)
                return;

            string textoComando = m_comandoPadrao;

            if (parameter is RibbonButton botaoRibbon
                && botaoRibbon.CommandParameter is string comandoRibbon
                && !string.IsNullOrWhiteSpace(comandoRibbon))
            {
                textoComando = comandoRibbon;
            }
            else if (parameter is string comandoTexto && !string.IsNullOrWhiteSpace(comandoTexto))
            {
                textoComando = comandoTexto;
            }

            string tipoParametro = parameter?.GetType().FullName ?? "(null)";
            documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Tipo de parâmetro no clique: {tipoParametro}");

            if (textoComando.StartsWith(m_prefixoComandoDummy, StringComparison.OrdinalIgnoreCase))
            {
                string mensagemDummy = textoComando[m_prefixoComandoDummy.Length..].Trim();
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] DUMMY acionado: {mensagemDummy}");
                return;
            }

            textoComando = NormalizarComando(textoComando);

            documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Enviando comando: {textoComando.Trim()}");
            documento.SendStringToExecute(textoComando, true, false, true);
        }

        private static string NormalizarComando(string textoComando)
        {
            string comando = textoComando.Trim();

            if (comando.Length > 0 && comando.Length % 2 == 0)
            {
                int metade = comando.Length / 2;
                string primeiraMetade = comando[..metade];
                string segundaMetade = comando[metade..];
                if (string.Equals(primeiraMetade, segundaMetade, StringComparison.Ordinal))
                    comando = primeiraMetade;
            }

            return comando + " ";
        }
    }
}