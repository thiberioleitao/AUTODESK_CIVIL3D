// RibbonInitializer.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using System;
using System.Linq;
using System.Runtime.Versioning;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using AplicacaoAutoCad = Autodesk.AutoCAD.ApplicationServices.Application;

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
        private const string m_panelFeatureLinesId = "ZAGO_TERRAPLENAGEM_FEATURE_LINES_PANEL";
        private const string m_prefixoComandoDummy = "DUMMY_PRINT:";
        private static bool m_aguardandoRibbon;

        public static void InicializarRibbon()
        {
            try
            {
                var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
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
                var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Erro em InitializeRibbon: {ex.Message}");
            }
        }

        private static void AoInicializarItem(object? remetente, RibbonItemEventArgs argumentos)
        {
            try
            {
                var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
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
                var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Erro em OnItemInitialized: {ex.Message}");
            }
        }

        private static void CriarRibbon(RibbonControl ribbon)
        {
            try
            {
                var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
                documento?.Editor.WriteMessage("\n[ZagoCivil3D] CreateRibbon chamado.");
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Total de abas atuais: {ribbon.Tabs.Count}");

                RibbonTab abaDrenagem = ObterOuCriarAba(ribbon, m_tabDrenagemId, "ZAGO - DRENAGEM");
                RibbonPanelSource painelCriarAlinhamentos = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarAlinhamentosId, "CRIAR - ALINHAMENTOS");
                AdicionarBotaoComando(
                    painelCriarAlinhamentos,
                    "ZAGO_CRIAR_ALINH_POR_POLI",
                    "Por Polilinha",
                    "ZAGO_CRIAR_ALINHAMENTOS_POR_POLILINHA ",
                    "AL",
                    "Cria alignments a partir das polilinhas de uma layer, sem ordenação específica.");
                AdicionarBotaoComando(
                    painelCriarAlinhamentos,
                    "ZAGO_CRIAR_ALINH_ORDENADOS",
                    "Ordenados\nH e V",
                    "ZAGO_CRIAR_ALINHAMENTOS_ORDENADOS ",
                    "AO",
                    "Cria alignments a partir de dois layers: primeiro as polilinhas horizontais (ordenadas Norte→Sul), em seguida as verticais (ordenadas Oeste→Leste), com numeração sequencial. Janela modeless.");
                AdicionarBotaoGrande(painelCriarAlinhamentos, "ZAGO_CRIAR_PONTOS_CRUZ", "Pontos nos\nCruzamentos", "CRIAR - ALINHAMENTOS > PONTOS NOS CRUZAMENTOS ENTRE ALINHAMENTOS", "PX");

                AdicionarBotaoGrande(painelCriarAlinhamentos, "ZAGO_CRIAR_PONTOS_CRUZ", "Pontos nos\nCruzamentos", "CRIAR - ALINHAMENTOS > PONTOS NOS CRUZAMENTOS ENTRE ALINHAMENTOS", "PX", "Cria pontos nos cruzamentos entre alignments (em definição).");

                RibbonPanelSource painelCriarPerfis = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarPerfisId, "CRIAR - PERFIS");
                AdicionarBotaoComando(painelCriarPerfis, "ZAGO_CRIAR_PERFIL_PROJ", "Perfis de\nProjeto", "ZAGO_CRIAR_PERFIS_DE_PROJETO ", "PP", "Cria perfis de projeto (layout profiles) para todos os alinhamentos a partir de uma superficie.");
                AdicionarBotaoComando(
                    painelCriarPerfis,
                    "ZAGO_CRIAR_PERFIL_TN",
                    "Perfis TN e\nTerraplenagem",
                    "ZAGO_CRIAR_PERFIS_TERRENO ",
                    "TN",
                    "Cria surface profiles (perfis do terreno natural e de terraplenagem) para todos os alinhamentos, a partir de uma superficie TIN. Janela modeless.");
                AdicionarBotaoComando(
                    painelCriarPerfis,
                    "ZAGO_CRIAR_PROFILE_VIEW",
                    "Profile\nView",
                    "ZAGO_CRIAR_PROFILE_VIEWS ",
                    "PV",
                    "Cria profile views para todos os alinhamentos, empilhados verticalmente a partir de uma coordenada inicial. Janela modeless.");

                RibbonPanelSource painelCriarCorredores = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarCorredoresId, "CRIAR - CORREDORES");
                AdicionarBotaoComando(
                    painelCriarCorredores,
                    "ZAGO_CRIAR_CORRED_POR_ALINH",
                    "Por\nAlinhamentos",
                    "ZAGO_CRIAR_CORREDORES_POR_ALINHAMENTOS ",
                    "CA",
                    "Cria um corredor vazio para cada alinhamento do desenho, com nome derivado do alinhamento (prefixo e sufixo opcionais). Janela modeless.");

                RibbonPanelSource painelCriarRegioes = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarRegioesId, "CRIAR - REGIOES");
                AdicionarBotaoComando(
                    painelCriarRegioes,
                    "ZAGO_CRIAR_REGIOES_CORREDORES",
                    "Regioes a partir\nde Corredores",
                    "ZAGO_CRIAR_REGIOES_CORREDORES ",
                    "RG",
                    "Quebra cada baseline de cada corredor em regioes, usando cruzamentos com outros alinhamentos, mudancas bruscas de direcao horizontal e mudancas de declividade no perfil como criterios. Apaga regioes existentes antes de recriar. Janela modeless.");

                RibbonPanelSource painelCriarCaixas = ObterOuCriarPainelFonte(abaDrenagem, m_panelCriarCaixasId, "CRIAR - CAIXAS");
                AdicionarBotaoGrande(painelCriarCaixas, "ZAGO_CRIAR_CAIXAS", "Criar\nCaixas", "CRIAR - CAIXAS > FUNCOES DE CRIACAO DE CAIXAS", "CX", "Funções de criação de caixas de drenagem (em definição).");

                RibbonPanelSource painelExportar = ObterOuCriarPainelFonte(abaDrenagem, m_panelExportarId, "EXPORTAR");
                AdicionarBotaoGrande(painelExportar, "ZAGO_EXPORTAR_CSV_BACIAS", "CSV Bacias e\nSubbacias", "EXPORTAR > CSV BACIAS/SUBBACIAS (ID, AREA, TALVEGUE, DECLIVIDADE, ID_JUSANTE)", "SB", "Exporta CSV com dados de bacias e subbacias (em definição).");
                AdicionarBotaoGrande(painelExportar, "ZAGO_EXPORTAR_CSV_CANAIS", "CSV Canais\ne Bueiros", "EXPORTAR > CSV CANAIS E BUEIROS", "CE", "Exporta CSV com dados de canais e bueiros (em definição).");

                RibbonPanelSource painelModificar = ObterOuCriarPainelFonte(abaDrenagem, m_panelModificarId, "MODIFICAR");
                AdicionarBotaoComando(
                    painelModificar,
                    "ZAGO_MUDAR_LABEL_SET_ALINH",
                    "Mudar Label Set\nAlinhamentos",
                    "ZAGO_MUDAR_LABEL_SET_ALINHAMENTOS ",
                    "LS",
                    "Aplica um Alignment Label Set Style a todos os alinhamentos do desenho, opcionalmente apagando os labels existentes antes. Janela modeless.");

                RibbonPanelSource painelAnotar = ObterOuCriarPainelFonte(abaDrenagem, m_panelAnotarId, "ANOTAR");
                AdicionarBotaoGrande(painelAnotar, "ZAGO_ANOTAR_LABELS_CORR", "Labels\nCorredores", "ANOTAR > ADICIONAR LABELS DOS TRECHOS/REGIOES DOS CORREDORES EM PLANTA", "LB", "Adiciona labels dos trechos/regiões dos corredores em planta (em definição).");

                RibbonPanelSource painelDeletar = ObterOuCriarPainelFonte(abaDrenagem, m_panelDeletarId, "DELETAR");
                AdicionarBotaoGrande(painelDeletar, "ZAGO_DELETAR_DUMMY", "Deletar\n(Dummy)", "DELETAR > FUNCOES EM DEFINICAO", "DL", "Funções de exclusão (em definição).");

                RibbonTab abaTerraplenagem = ObterOuCriarAba(ribbon, m_tabTerraplenagemId, "ZAGO - TERRAPLENAGEM");
                RibbonPanelSource painelFeatureLinesTerraplenagem = ObterOuCriarPainelFonte(
                    abaTerraplenagem,
                    m_panelFeatureLinesId,
                    "FEATURE LINES");
                AdicionarBotaoComando(
                    painelFeatureLinesTerraplenagem,
                    "ZAGO_TERRAPL_FEATURE_LINES_SEPARADAS",
                    "Feature Lines\nSeparadas",
                    "ZAGO_TERRAPLENAGEM_FEATURE_LINES_SEPARADAS ",
                    "TP",
                    "Cria feature lines separadas a partir de polilinhas de terraplenagem.");
                AdicionarBotaoComando(
                    painelFeatureLinesTerraplenagem,
                    "ZAGO_TERRAPL_AJUSTAR_DEFLEXAO",
                    "Ajustar\nDeflexao",
                    "ZAGO_AJUSTAR_DEFLEXAO_FEATURE_LINES ",
                    "DF",
                    "Ajusta iterativamente as cotas dos PI points das feature lines ate que a deflexao entre segmentos adjacentes fique dentro do limite. Duas passadas (montante->jusante e jusante->montante). Janela modeless.");

                documento?.Editor.WriteMessage("\n[ZagoCivil3D] Aba e botão criados com sucesso.");
                documento?.Editor.WriteMessage($"\n[ZagoCivil3D] Aba ativa: {ribbon.ActiveTab?.Id ?? "(nenhuma)"}");
            }
            catch (Exception ex)
            {
                var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
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
            string siglaIcone,
            string descricao = "")
        {
            AdicionarBotao(fontePainel, idBotao, textoBotao, m_prefixoComandoDummy + mensagemDummy, siglaIcone, descricao);
        }

        private static void AdicionarBotaoComando(
            RibbonPanelSource fontePainel,
            string idBotao,
            string textoBotao,
            string nomeComando,
            string siglaIcone,
            string descricao = "")
        {
            AdicionarBotao(fontePainel, idBotao, textoBotao, nomeComando, siglaIcone, descricao);
        }

        private static void AdicionarBotao(
            RibbonPanelSource fontePainel,
            string idBotao,
            string textoBotao,
            string parametroComando,
            string siglaIcone,
            string descricao = "")
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
                ShowImage = true,
                Size = RibbonItemSize.Large,
                Orientation = System.Windows.Controls.Orientation.Vertical,
                LargeImage = CriarIcone(siglaIcone, grande: true),
                Image = CriarIcone(siglaIcone, grande: false),
                CommandParameter = parametroComando,
                CommandHandler = new ManipuladorComandoRibbon()
            };

            if (!string.IsNullOrWhiteSpace(descricao))
            {
                botao.Description = descricao;
                botao.ToolTip = new RibbonToolTip
                {
                    Title = textoBotao.Replace("\n", " "),
                    Content = descricao,
                    Command = idBotao
                };
            }

            fontePainel.Items.Add(botao);
        }

        /// <summary>
        /// Gera um ícone vetorial representativo do comando, no estilo dos ícones
        /// nativos do Civil 3D: fundo transparente, glifo desenhado com a cor de
        /// destaque da categoria. Quando <paramref name="grande"/> é verdadeiro o
        /// ícone é 32x32 (LargeImage), caso contrário 16x16 (Image).
        /// </summary>
        private static ImageSource CriarIcone(string sigla, bool grande)
        {
            int tamanho = grande ? 32 : 16;
            var visual = new DrawingVisual();
            using (DrawingContext contexto = visual.RenderOpen())
            {
                Color corPrincipal = ObterCorIcone(sigla);
                var pincelPrincipal = new SolidColorBrush(corPrincipal);
                pincelPrincipal.Freeze();
                var pincelClaro = new SolidColorBrush(Color.FromArgb(60, corPrincipal.R, corPrincipal.G, corPrincipal.B));
                pincelClaro.Freeze();
                var pincelEscuro = new SolidColorBrush(Color.FromRgb(
                    (byte)(corPrincipal.R * 0.55),
                    (byte)(corPrincipal.G * 0.55),
                    (byte)(corPrincipal.B * 0.55)));
                pincelEscuro.Freeze();
                var pincelCinza = new SolidColorBrush(Color.FromRgb(140, 140, 140));
                pincelCinza.Freeze();

                DesenharGlifo(contexto, sigla, tamanho, pincelPrincipal, pincelClaro, pincelEscuro, pincelCinza);
            }

            var bitmap = new RenderTargetBitmap(tamanho, tamanho, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(visual);
            bitmap.Freeze();
            return bitmap;
        }

        /// <summary>
        /// Despachante que escolhe o glifo apropriado para cada sigla.
        /// Todos os desenhos usam coordenadas normalizadas (0..tamanho)
        /// para funcionar em 16x16 e 32x32.
        /// </summary>
        private static void DesenharGlifo(
            DrawingContext g,
            string sigla,
            int tamanho,
            Brush corPrincipal,
            Brush corClara,
            Brush corEscura,
            Brush corCinza)
        {
            double s = tamanho;
            double linhaFina = tamanho <= 16 ? 1.2 : 1.8;
            double linhaGrossa = tamanho <= 16 ? 1.6 : 2.4;
            var penPrincipal = new Pen(corPrincipal, linhaGrossa) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round, LineJoin = PenLineJoin.Round };
            var penEscuro = new Pen(corEscura, linhaFina) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round, LineJoin = PenLineJoin.Round };
            var penCinza = new Pen(corCinza, linhaFina) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round };
            penPrincipal.Freeze();
            penEscuro.Freeze();
            penCinza.Freeze();

            switch (sigla)
            {
                case "AL":
                    // Polilinha com vértices (símbolo de alignment genérico)
                    {
                        var polilinha = new StreamGeometry();
                        using (var ctx = polilinha.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.12, s * 0.75), false, false);
                            ctx.LineTo(new Point(s * 0.35, s * 0.45), true, false);
                            ctx.LineTo(new Point(s * 0.60, s * 0.60), true, false);
                            ctx.LineTo(new Point(s * 0.88, s * 0.22), true, false);
                        }
                        polilinha.Freeze();
                        g.DrawGeometry(null, penPrincipal, polilinha);
                        double r = s * 0.08;
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.12, s * 0.75), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.35, s * 0.45), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.60, s * 0.60), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.88, s * 0.22), r, r);
                    }
                    break;

                case "AO":
                    // Grade com setas indicando ordem N→S (vertical) e O→L (horizontal)
                    {
                        // Linhas horizontais paralelas (grupo de horizontais)
                        g.DrawLine(penEscuro, new Point(s * 0.15, s * 0.30), new Point(s * 0.85, s * 0.30));
                        g.DrawLine(penEscuro, new Point(s * 0.15, s * 0.50), new Point(s * 0.85, s * 0.50));
                        g.DrawLine(penEscuro, new Point(s * 0.15, s * 0.70), new Point(s * 0.85, s * 0.70));
                        // Seta vertical N→S à esquerda
                        g.DrawLine(penPrincipal, new Point(s * 0.08, s * 0.15), new Point(s * 0.08, s * 0.85));
                        var setaBaixo = new StreamGeometry();
                        using (var ctx = setaBaixo.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.08, s * 0.90), true, true);
                            ctx.LineTo(new Point(s * 0.02, s * 0.78), true, false);
                            ctx.LineTo(new Point(s * 0.14, s * 0.78), true, false);
                        }
                        setaBaixo.Freeze();
                        g.DrawGeometry(corPrincipal, null, setaBaixo);
                        // Seta horizontal O→L no topo
                        g.DrawLine(penPrincipal, new Point(s * 0.20, s * 0.10), new Point(s * 0.90, s * 0.10));
                        var setaDireita = new StreamGeometry();
                        using (var ctx = setaDireita.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.95, s * 0.10), true, true);
                            ctx.LineTo(new Point(s * 0.83, s * 0.04), true, false);
                            ctx.LineTo(new Point(s * 0.83, s * 0.16), true, false);
                        }
                        setaDireita.Freeze();
                        g.DrawGeometry(corPrincipal, null, setaDireita);
                    }
                    break;

                case "PX":
                    // Duas linhas cruzando com ponto no cruzamento
                    g.DrawLine(penEscuro, new Point(s * 0.10, s * 0.10), new Point(s * 0.90, s * 0.90));
                    g.DrawLine(penEscuro, new Point(s * 0.90, s * 0.10), new Point(s * 0.10, s * 0.90));
                    g.DrawEllipse(corPrincipal, null, new Point(s * 0.50, s * 0.50), s * 0.14, s * 0.14);
                    break;

                case "PP":
                    // Perfil longitudinal: linha base + curva de perfil acima
                    g.DrawLine(penCinza, new Point(s * 0.10, s * 0.80), new Point(s * 0.90, s * 0.80));
                    {
                        var curva = new StreamGeometry();
                        using (var ctx = curva.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.10, s * 0.60), false, false);
                            ctx.LineTo(new Point(s * 0.30, s * 0.30), true, false);
                            ctx.LineTo(new Point(s * 0.55, s * 0.50), true, false);
                            ctx.LineTo(new Point(s * 0.75, s * 0.25), true, false);
                            ctx.LineTo(new Point(s * 0.90, s * 0.45), true, false);
                        }
                        curva.Freeze();
                        g.DrawGeometry(null, penPrincipal, curva);
                    }
                    break;

                case "TN":
                    // Perfil do terreno natural: zigzag de montanha
                    {
                        var montanha = new StreamGeometry();
                        using (var ctx = montanha.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.08, s * 0.80), true, true);
                            ctx.LineTo(new Point(s * 0.28, s * 0.40), true, false);
                            ctx.LineTo(new Point(s * 0.48, s * 0.60), true, false);
                            ctx.LineTo(new Point(s * 0.70, s * 0.25), true, false);
                            ctx.LineTo(new Point(s * 0.92, s * 0.55), true, false);
                            ctx.LineTo(new Point(s * 0.92, s * 0.80), true, false);
                        }
                        montanha.Freeze();
                        g.DrawGeometry(corClara, penPrincipal, montanha);
                        g.DrawLine(penCinza, new Point(s * 0.05, s * 0.80), new Point(s * 0.95, s * 0.80));
                    }
                    break;

                case "PV":
                    // Profile view: eixos (L + curva)
                    g.DrawLine(penEscuro, new Point(s * 0.15, s * 0.15), new Point(s * 0.15, s * 0.85));
                    g.DrawLine(penEscuro, new Point(s * 0.15, s * 0.85), new Point(s * 0.90, s * 0.85));
                    {
                        var curva = new StreamGeometry();
                        using (var ctx = curva.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.20, s * 0.70), false, false);
                            ctx.LineTo(new Point(s * 0.40, s * 0.40), true, false);
                            ctx.LineTo(new Point(s * 0.62, s * 0.55), true, false);
                            ctx.LineTo(new Point(s * 0.85, s * 0.30), true, false);
                        }
                        curva.Freeze();
                        g.DrawGeometry(null, penPrincipal, curva);
                    }
                    break;

                case "CA":
                    // Corredor a partir de alinhamento: polilinha (eixo) com pontos
                    // nos vertices e uma linha paralela de corredor por baixo.
                    {
                        // Linha paralela inferior (faixa do corredor)
                        var paralela = new StreamGeometry();
                        using (var ctx = paralela.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.12, s * 0.80), false, false);
                            ctx.LineTo(new Point(s * 0.35, s * 0.55), true, false);
                            ctx.LineTo(new Point(s * 0.60, s * 0.70), true, false);
                            ctx.LineTo(new Point(s * 0.88, s * 0.32), true, false);
                        }
                        paralela.Freeze();
                        g.DrawGeometry(null, penEscuro, paralela);

                        // Linha principal (alinhamento / eixo do corredor)
                        var eixo = new StreamGeometry();
                        using (var ctx = eixo.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.12, s * 0.60), false, false);
                            ctx.LineTo(new Point(s * 0.35, s * 0.35), true, false);
                            ctx.LineTo(new Point(s * 0.60, s * 0.50), true, false);
                            ctx.LineTo(new Point(s * 0.88, s * 0.12), true, false);
                        }
                        eixo.Freeze();
                        g.DrawGeometry(null, penPrincipal, eixo);

                        // Pontos nos vertices do alinhamento
                        double r = s * 0.075;
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.12, s * 0.60), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.35, s * 0.35), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.60, s * 0.50), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.88, s * 0.12), r, r);
                    }
                    break;

                case "CO":
                    // Corredor: duas curvas paralelas com hachura central
                    {
                        var curvaA = new StreamGeometry();
                        using (var ctx = curvaA.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.10, s * 0.30), false, false);
                            ctx.LineTo(new Point(s * 0.90, s * 0.30), true, false);
                        }
                        curvaA.Freeze();
                        var curvaB = new StreamGeometry();
                        using (var ctx = curvaB.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.10, s * 0.70), false, false);
                            ctx.LineTo(new Point(s * 0.90, s * 0.70), true, false);
                        }
                        curvaB.Freeze();
                        g.DrawGeometry(null, penPrincipal, curvaA);
                        g.DrawGeometry(null, penPrincipal, curvaB);
                        // traços centrais (faixa)
                        double y = s * 0.50;
                        g.DrawLine(penCinza, new Point(s * 0.18, y), new Point(s * 0.30, y));
                        g.DrawLine(penCinza, new Point(s * 0.44, y), new Point(s * 0.56, y));
                        g.DrawLine(penCinza, new Point(s * 0.70, y), new Point(s * 0.82, y));
                    }
                    break;

                case "RG":
                    // Região: polígono fechado preenchido
                    {
                        var poli = new StreamGeometry();
                        using (var ctx = poli.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.20, s * 0.20), true, true);
                            ctx.LineTo(new Point(s * 0.80, s * 0.25), true, false);
                            ctx.LineTo(new Point(s * 0.85, s * 0.70), true, false);
                            ctx.LineTo(new Point(s * 0.50, s * 0.88), true, false);
                            ctx.LineTo(new Point(s * 0.18, s * 0.65), true, false);
                        }
                        poli.Freeze();
                        g.DrawGeometry(corClara, penPrincipal, poli);
                    }
                    break;

                case "CX":
                    // Caixa (catch basin): retângulo com tubulação saindo
                    g.DrawRectangle(corClara, penPrincipal, new Rect(s * 0.22, s * 0.30, s * 0.42, s * 0.42));
                    g.DrawRectangle(corPrincipal, null, new Rect(s * 0.22, s * 0.30, s * 0.42, s * 0.08));
                    g.DrawLine(penEscuro, new Point(s * 0.64, s * 0.55), new Point(s * 0.90, s * 0.55));
                    break;

                case "SB":
                    // Bacias: gotas / curvas de nível
                    {
                        var c1 = new StreamGeometry();
                        using (var ctx = c1.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.20, s * 0.30), false, false);
                            ctx.BezierTo(new Point(s * 0.40, s * 0.20), new Point(s * 0.60, s * 0.20), new Point(s * 0.80, s * 0.35), true, false);
                        }
                        c1.Freeze();
                        var c2 = new StreamGeometry();
                        using (var ctx = c2.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.15, s * 0.50), false, false);
                            ctx.BezierTo(new Point(s * 0.40, s * 0.40), new Point(s * 0.60, s * 0.40), new Point(s * 0.85, s * 0.55), true, false);
                        }
                        c2.Freeze();
                        var c3 = new StreamGeometry();
                        using (var ctx = c3.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.10, s * 0.72), false, false);
                            ctx.BezierTo(new Point(s * 0.40, s * 0.62), new Point(s * 0.60, s * 0.62), new Point(s * 0.90, s * 0.78), true, false);
                        }
                        c3.Freeze();
                        g.DrawGeometry(null, penPrincipal, c1);
                        g.DrawGeometry(null, penEscuro, c2);
                        g.DrawGeometry(null, penCinza, c3);
                    }
                    break;

                case "CE":
                    // Canais: tubo em perspectiva
                    g.DrawEllipse(corClara, penPrincipal, new Point(s * 0.25, s * 0.50), s * 0.12, s * 0.25);
                    g.DrawRectangle(corClara, null, new Rect(s * 0.25, s * 0.25, s * 0.55, s * 0.50));
                    g.DrawLine(penPrincipal, new Point(s * 0.25, s * 0.25), new Point(s * 0.80, s * 0.25));
                    g.DrawLine(penPrincipal, new Point(s * 0.25, s * 0.75), new Point(s * 0.80, s * 0.75));
                    g.DrawEllipse(null, penPrincipal, new Point(s * 0.80, s * 0.50), s * 0.12, s * 0.25);
                    break;

                case "MD":
                    // Modificar: chave/ferramenta
                    {
                        var chave = new StreamGeometry();
                        using (var ctx = chave.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.70, s * 0.15), true, true);
                            ctx.LineTo(new Point(s * 0.88, s * 0.32), true, false);
                            ctx.LineTo(new Point(s * 0.72, s * 0.48), true, false);
                            ctx.LineTo(new Point(s * 0.60, s * 0.36), true, false);
                            ctx.LineTo(new Point(s * 0.25, s * 0.72), true, false);
                            ctx.LineTo(new Point(s * 0.12, s * 0.85), true, false);
                            ctx.LineTo(new Point(s * 0.22, s * 0.75), true, false);
                            ctx.LineTo(new Point(s * 0.58, s * 0.40), true, false);
                            ctx.LineTo(new Point(s * 0.46, s * 0.28), true, false);
                            ctx.LineTo(new Point(s * 0.62, s * 0.12), true, false);
                        }
                        chave.Freeze();
                        g.DrawGeometry(corPrincipal, null, chave);
                    }
                    break;

                case "LB":
                    // Label: etiqueta com linha de chamada
                    g.DrawLine(penCinza, new Point(s * 0.10, s * 0.80), new Point(s * 0.40, s * 0.55));
                    g.DrawRectangle(corPrincipal, null, new Rect(s * 0.40, s * 0.30, s * 0.48, s * 0.30));
                    g.DrawLine(penEscuro, new Point(s * 0.50, s * 0.42), new Point(s * 0.82, s * 0.42));
                    g.DrawLine(penEscuro, new Point(s * 0.50, s * 0.50), new Point(s * 0.76, s * 0.50));
                    break;

                case "DL":
                    // Deletar: lixeira
                    g.DrawRectangle(corClara, penPrincipal, new Rect(s * 0.25, s * 0.32, s * 0.50, s * 0.52));
                    g.DrawRectangle(corPrincipal, null, new Rect(s * 0.20, s * 0.22, s * 0.60, s * 0.10));
                    g.DrawLine(penEscuro, new Point(s * 0.38, s * 0.42), new Point(s * 0.38, s * 0.76));
                    g.DrawLine(penEscuro, new Point(s * 0.50, s * 0.42), new Point(s * 0.50, s * 0.76));
                    g.DrawLine(penEscuro, new Point(s * 0.62, s * 0.42), new Point(s * 0.62, s * 0.76));
                    break;

                case "LS":
                    // Label Set: polilinha de alignment com etiquetas de estacao
                    {
                        // Alinhamento base (mesma silhueta do AL para coerencia visual)
                        var polilinha = new StreamGeometry();
                        using (var ctx = polilinha.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.10, s * 0.70), false, false);
                            ctx.LineTo(new Point(s * 0.35, s * 0.50), true, false);
                            ctx.LineTo(new Point(s * 0.60, s * 0.60), true, false);
                            ctx.LineTo(new Point(s * 0.90, s * 0.30), true, false);
                        }
                        polilinha.Freeze();
                        g.DrawGeometry(null, penEscuro, polilinha);

                        // Tick marks (estacoes) perpendiculares ao alinhamento
                        g.DrawLine(penPrincipal, new Point(s * 0.22, s * 0.54), new Point(s * 0.30, s * 0.70));
                        g.DrawLine(penPrincipal, new Point(s * 0.48, s * 0.49), new Point(s * 0.50, s * 0.67));
                        g.DrawLine(penPrincipal, new Point(s * 0.72, s * 0.40), new Point(s * 0.80, s * 0.56));

                        // Etiqueta (caixa) representando o label
                        g.DrawRectangle(corPrincipal, null, new Rect(s * 0.55, s * 0.08, s * 0.35, s * 0.16));
                        g.DrawLine(penCinza, new Point(s * 0.60, s * 0.16), new Point(s * 0.82, s * 0.16));
                    }
                    break;

                case "TP":
                    // Feature line: linha ondulada com pontos
                    {
                        var onda = new StreamGeometry();
                        using (var ctx = onda.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.10, s * 0.60), false, false);
                            ctx.BezierTo(new Point(s * 0.30, s * 0.20), new Point(s * 0.50, s * 0.80), new Point(s * 0.90, s * 0.40), true, false);
                        }
                        onda.Freeze();
                        g.DrawGeometry(null, penPrincipal, onda);
                        double r = s * 0.08;
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.10, s * 0.60), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.90, s * 0.40), r, r);
                    }
                    break;

                case "DF":
                    // Deflexao: tres pontos com segmentos formando angulo no ponto central.
                    // O vertice central aparece destacado em corPrincipal (o ponto ajustado).
                    {
                        var pontoEsq = new Point(s * 0.12, s * 0.70);
                        var pontoMeio = new Point(s * 0.50, s * 0.30);
                        var pontoDir = new Point(s * 0.88, s * 0.62);

                        // Segmentos adjacentes que formam a deflexao
                        var segmentos = new StreamGeometry();
                        using (var ctx = segmentos.Open())
                        {
                            ctx.BeginFigure(pontoEsq, false, false);
                            ctx.LineTo(pontoMeio, true, false);
                            ctx.LineTo(pontoDir, true, false);
                        }
                        segmentos.Freeze();
                        g.DrawGeometry(null, penPrincipal, segmentos);

                        // Linha horizontal pontilhada indicando "sem deflexao"
                        var penTracejada = new Pen(corCinza, linhaFina) { DashStyle = new DashStyle(new double[] { 2, 2 }, 0) };
                        penTracejada.Freeze();
                        g.DrawLine(penTracejada, new Point(s * 0.12, s * 0.30), new Point(s * 0.88, s * 0.30));

                        // Pontos nos extremos e, destacado, no vertice central
                        double rLateral = s * 0.06;
                        double rCentro = s * 0.10;
                        g.DrawEllipse(corEscura, null, pontoEsq, rLateral, rLateral);
                        g.DrawEllipse(corEscura, null, pontoDir, rLateral, rLateral);
                        g.DrawEllipse(corPrincipal, null, pontoMeio, rCentro, rCentro);
                    }
                    break;

                default:
                    // Fallback: disco simples
                    g.DrawEllipse(corPrincipal, null, new Point(s * 0.5, s * 0.5), s * 0.3, s * 0.3);
                    break;
            }
        }

        private static Color ObterCorIcone(string sigla)
        {
            return sigla switch
            {
                "AL" => Color.FromRgb(0, 120, 212),
                "AO" => Color.FromRgb(0, 90, 180),
                "PX" => Color.FromRgb(90, 90, 220),
                "PP" => Color.FromRgb(0, 153, 102),
                "TN" => Color.FromRgb(46, 139, 87),
                "PV" => Color.FromRgb(0, 153, 153),
                "CO" => Color.FromRgb(128, 0, 128),
                "CA" => Color.FromRgb(128, 0, 128),
                "RG" => Color.FromRgb(184, 134, 11),
                "CX" => Color.FromRgb(210, 105, 30),
                "SB" => Color.FromRgb(220, 20, 60),
                "CE" => Color.FromRgb(178, 34, 34),
                "MD" => Color.FromRgb(255, 140, 0),
                "LS" => Color.FromRgb(255, 140, 0),
                "LB" => Color.FromRgb(72, 61, 139),
                "DL" => Color.FromRgb(169, 0, 0),
                "TP" => Color.FromRgb(105, 105, 105),
                "DF" => Color.FromRgb(180, 110, 40),
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
            var documento = AplicacaoAutoCad.DocumentManager.MdiActiveDocument;
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
