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
        private const string m_tabId = "ZAGO_TAB";
        private const string m_panelAlinhamentosId = "ZAGO_ALINHAMENTOS_PANEL";
        private const string m_panelPerfisId = "ZAGO_PERFIS_PANEL";
        private const string m_panelProfileViewsId = "ZAGO_PROFILE_VIEWS_PANEL";
        private const string m_panelCogopointsId = "ZAGO_COGOPOINTS_PANEL";
        private const string m_panelCorredoresId = "ZAGO_CORREDORES_PANEL";
        private const string m_panelFeatureLinesId = "ZAGO_FEATURE_LINES_PANEL";
        private const string m_panelBaciasId = "ZAGO_BACIAS_PANEL";
        private const string m_panelCaixasId = "ZAGO_CAIXAS_PANEL";
        private const string m_panelTopologiaId = "ZAGO_TOPOLOGIA_PANEL";
        private const string m_panelDeletarId = "ZAGO_DELETAR_PANEL";
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

                RibbonTab aba = ObterOuCriarAba(ribbon, m_tabId, "ZAGO");

                // ALINHAMENTOS (dois dropdowns lado-a-lado: Criar e Anotação)
                RibbonPanelSource painelAlinhamentos = ObterOuCriarPainelFonte(aba, m_panelAlinhamentosId, "ALINHAMENTOS");

                RibbonSplitButton dropdownCriarAlinh = CriarDropdown(
                    painelAlinhamentos,
                    "ZAGO_DROPDOWN_ALINH_CRIAR",
                    "Criar",
                    "AL",
                    "Comandos de criação de alinhamentos.");
                AdicionarItemDropdown(
                    dropdownCriarAlinh,
                    "ZAGO_CRIAR_ALINH_POR_POLI",
                    "Por Polilinha",
                    "ZAGO_CRIAR_ALINHAMENTOS_POR_POLILINHA ",
                    "AL",
                    "Cria alignments a partir das polilinhas de uma layer, sem ordenação específica.");
                AdicionarItemDropdown(
                    dropdownCriarAlinh,
                    "ZAGO_CRIAR_ALINH_ORDENADOS",
                    "Ordenados H e V",
                    "ZAGO_CRIAR_ALINHAMENTOS_ORDENADOS ",
                    "AO",
                    "Cria alignments a partir de dois layers: primeiro as polilinhas horizontais (ordenadas Norte→Sul), em seguida as verticais (ordenadas Oeste→Leste), com numeração sequencial. Janela modeless.");

                RibbonSplitButton dropdownAnotarAlinh = CriarDropdown(
                    painelAlinhamentos,
                    "ZAGO_DROPDOWN_ALINH_ANOTAR",
                    "Anotação",
                    "LS",
                    "Comandos de anotação de alinhamentos.");
                AdicionarItemDropdown(
                    dropdownAnotarAlinh,
                    "ZAGO_MUDAR_LABEL_SET_ALINH",
                    "Mudar Label Set",
                    "ZAGO_MUDAR_LABEL_SET_ALINHAMENTOS ",
                    "LS",
                    "Aplica um Alignment Label Set Style a todos os alinhamentos do desenho, opcionalmente apagando os labels existentes antes. Janela modeless.");

                RibbonSplitButton dropdownConverterAlinh = CriarDropdown(
                    painelAlinhamentos,
                    "ZAGO_DROPDOWN_ALINH_CONVERTER",
                    "Converter",
                    "AP",
                    "Comandos de conversao de alinhamentos em outras entidades.");
                AdicionarItemDropdown(
                    dropdownConverterAlinh,
                    "ZAGO_CONVERTER_ALINH_EM_POLI",
                    "Em Polilinhas 2D",
                    "ZAGO_CONVERTER_ALINHAMENTOS_EM_POLILINHAS ",
                    "AP",
                    "Converte todos os alinhamentos do desenho em polilinhas 2D (Polyline). Arcos circulares viram vertices com bulge; espirais sao discretizadas em segmentos retos pelo passo configurado. Janela modeless.");

                // PERFIS (Criar)
                RibbonPanelSource painelPerfis = ObterOuCriarPainelFonte(aba, m_panelPerfisId, "PERFIS");
                RibbonSplitButton dropdownCriarPerfis = CriarDropdown(
                    painelPerfis,
                    "ZAGO_DROPDOWN_PERFIS_CRIAR",
                    "Criar",
                    "PP",
                    "Comandos de criação de perfis.");
                AdicionarItemDropdown(
                    dropdownCriarPerfis,
                    "ZAGO_CRIAR_PERFIL_PROJ",
                    "Perfis de Projeto",
                    "ZAGO_CRIAR_PERFIS_DE_PROJETO ",
                    "PP",
                    "Cria perfis de projeto (layout profiles) para todos os alinhamentos a partir de uma superficie.");
                AdicionarItemDropdown(
                    dropdownCriarPerfis,
                    "ZAGO_CRIAR_PERFIL_TN",
                    "Perfis de superfície",
                    "ZAGO_CRIAR_PERFIS_TERRENO ",
                    "TN",
                    "Cria surface profiles (perfis do terreno natural e de terraplenagem) para todos os alinhamentos, a partir de uma superficie TIN. Janela modeless.");

                // PROFILE VIEWS (Criar)
                RibbonPanelSource painelProfileViews = ObterOuCriarPainelFonte(aba, m_panelProfileViewsId, "PROFILE VIEWS");
                RibbonSplitButton dropdownCriarProfileViews = CriarDropdown(
                    painelProfileViews,
                    "ZAGO_DROPDOWN_PROFILEVIEWS_CRIAR",
                    "Criar",
                    "PV",
                    "Comandos de criação de profile views.");
                AdicionarItemDropdown(
                    dropdownCriarProfileViews,
                    "ZAGO_CRIAR_PROFILE_VIEW",
                    "Profile View",
                    "ZAGO_CRIAR_PROFILE_VIEWS ",
                    "PV",
                    "Cria profile views para todos os alinhamentos, empilhados verticalmente a partir de uma coordenada inicial. Janela modeless.");

                // COGOPOINTS (Criar)
                RibbonPanelSource painelCogopoints = ObterOuCriarPainelFonte(aba, m_panelCogopointsId, "COGOPOINTS");
                RibbonSplitButton dropdownCriarCogopoints = CriarDropdown(
                    painelCogopoints,
                    "ZAGO_DROPDOWN_COGOPOINTS_CRIAR",
                    "Criar",
                    "PX",
                    "Comandos de criação de cogopoints.");
                AdicionarItemDropdown(
                    dropdownCriarCogopoints,
                    "ZAGO_CRIAR_PONTOS_CRUZ",
                    "Pontos nos Cruzamentos",
                    "ZAGO_CRIAR_PONTOS_CRUZAMENTOS_ALINHAMENTOS ",
                    "PX",
                    "Cria CogoPoints nos cruzamentos entre alinhamentos, organizados em blocos (trackers) com rotulos sequenciais. Janela modeless.");

                // CORREDORES (Criar + Anotação)
                RibbonPanelSource painelCorredores = ObterOuCriarPainelFonte(aba, m_panelCorredoresId, "CORREDORES");
                RibbonSplitButton dropdownCriarCorredores = CriarDropdown(
                    painelCorredores,
                    "ZAGO_DROPDOWN_CORREDORES_CRIAR",
                    "Criar",
                    "CO",
                    "Comandos de criação de corredores e regiões.");
                AdicionarItemDropdown(
                    dropdownCriarCorredores,
                    "ZAGO_CRIAR_CORRED_POR_ALINH",
                    "Por Alinhamentos",
                    "ZAGO_CRIAR_CORREDORES_POR_ALINHAMENTOS ",
                    "CA",
                    "Cria um corredor vazio para cada alinhamento do desenho, com nome derivado do alinhamento (prefixo e sufixo opcionais). Janela modeless.");
                AdicionarItemDropdown(
                    dropdownCriarCorredores,
                    "ZAGO_CRIAR_REGIOES_CORREDORES",
                    "Regiões a partir de Corredores",
                    "ZAGO_CRIAR_REGIOES_CORREDORES ",
                    "RG",
                    "Quebra cada baseline de cada corredor em regioes, usando cruzamentos com outros alinhamentos, mudancas bruscas de direcao horizontal e mudancas de declividade no perfil como criterios. Apaga regioes existentes antes de recriar. Janela modeless.");

                RibbonSplitButton dropdownAnotarCorredores = CriarDropdown(
                    painelCorredores,
                    "ZAGO_DROPDOWN_CORREDORES_ANOTAR",
                    "Anotação",
                    "LB",
                    "Comandos de anotação de corredores.");
                AdicionarItemDropdownDummy(
                    dropdownAnotarCorredores,
                    "ZAGO_ANOTAR_LABELS_CORR",
                    "Labels Corredores",
                    "CORREDORES > ADICIONAR LABELS DOS TRECHOS/REGIOES DOS CORREDORES EM PLANTA",
                    "LB",
                    "Adiciona labels dos trechos/regiões dos corredores em planta (em definição).");

                // FEATURE LINES (Criar + Modificar)
                RibbonPanelSource painelFeatureLines = ObterOuCriarPainelFonte(aba, m_panelFeatureLinesId, "FEATURE LINES");
                RibbonSplitButton dropdownCriarFeatureLines = CriarDropdown(
                    painelFeatureLines,
                    "ZAGO_DROPDOWN_FEATURELINES_CRIAR",
                    "Criar",
                    "TP",
                    "Comandos de criação de feature lines.");
                AdicionarItemDropdown(
                    dropdownCriarFeatureLines,
                    "ZAGO_TERRAPL_FEATURE_LINES_SEPARADAS",
                    "Feature Lines Separadas",
                    "ZAGO_TERRAPLENAGEM_FEATURE_LINES_SEPARADAS ",
                    "TP",
                    "Cria feature lines separadas a partir de polilinhas de terraplenagem.");

                RibbonSplitButton dropdownModificarFeatureLines = CriarDropdown(
                    painelFeatureLines,
                    "ZAGO_DROPDOWN_FEATURELINES_MODIFICAR",
                    "Modificar",
                    "DF",
                    "Comandos de modificação de feature lines.");
                AdicionarItemDropdown(
                    dropdownModificarFeatureLines,
                    "ZAGO_TERRAPL_AJUSTAR_DEFLEXAO",
                    "Ajustar Deflexão",
                    "ZAGO_AJUSTAR_DEFLEXAO_FEATURE_LINES ",
                    "DF",
                    "Ajusta iterativamente as cotas dos PI points das feature lines ate que a deflexao entre segmentos adjacentes fique dentro do limite. Duas passadas (montante->jusante e jusante->montante). Janela modeless.");

                // BACIAS (Criar + Exportar)
                RibbonPanelSource painelBacias = ObterOuCriarPainelFonte(aba, m_panelBaciasId, "BACIAS");
                RibbonSplitButton dropdownCriarBacias = CriarDropdown(
                    painelBacias,
                    "ZAGO_DROPDOWN_BACIAS_CRIAR",
                    "Criar",
                    "BC",
                    "Comandos de criação de bacias e talvegues.");
                AdicionarItemDropdown(
                    dropdownCriarBacias,
                    "ZAGO_CRIAR_CATCHMENTS_HATCHS",
                    "Catchments de Hatches",
                    "ZAGO_CRIAR_CATCHMENTS_DE_HATCHS ",
                    "BC",
                    "Cria catchments (subbacias) a partir das hatches de uma layer, associando o MText de ID e a polilinha de talvegue contidos em cada hatch. O talvegue vira flow path do catchment. Janela modeless.");
                AdicionarItemDropdown(
                    dropdownCriarBacias,
                    "ZAGO_CRIAR_TALVEGUES_CATCH",
                    "Talvegues dos Catchments",
                    "ZAGO_CRIAR_TALVEGUES_CATCHMENTS ",
                    "TC",
                    "Define o FlowPath dos catchments existentes a partir das polilinhas desenhadas em uma layer (padrao ZAGO: HDR-TALVEGUES SUBBACIAS). Cada polyline vira o talvegue do catchment que a contem. Janela modeless.");

                RibbonSplitButton dropdownExportarBacias = CriarDropdown(
                    painelBacias,
                    "ZAGO_DROPDOWN_BACIAS_EXPORTAR",
                    "Exportar",
                    "SB",
                    "Comandos de exportação de dados de bacias.");
                AdicionarItemDropdownDummy(
                    dropdownExportarBacias,
                    "ZAGO_EXPORTAR_CSV_BACIAS",
                    "CSV Bacias e Subbacias",
                    "BACIAS > CSV BACIAS/SUBBACIAS (ID, AREA, TALVEGUE, DECLIVIDADE, ID_JUSANTE)",
                    "SB",
                    "Exporta CSV com dados de bacias e subbacias (em definição).");

                // CAIXAS (Criar + Exportar)
                RibbonPanelSource painelCaixas = ObterOuCriarPainelFonte(aba, m_panelCaixasId, "CAIXAS");
                RibbonSplitButton dropdownCriarCaixas = CriarDropdown(
                    painelCaixas,
                    "ZAGO_DROPDOWN_CAIXAS_CRIAR",
                    "Criar",
                    "CX",
                    "Comandos de criação de caixas de drenagem.");
                AdicionarItemDropdownDummy(
                    dropdownCriarCaixas,
                    "ZAGO_CRIAR_CAIXAS",
                    "Criar Caixas",
                    "CAIXAS > FUNCOES DE CRIACAO DE CAIXAS",
                    "CX",
                    "Funções de criação de caixas de drenagem (em definição).");

                // TOPOLOGIA (Exportar)
                RibbonPanelSource painelTopologia = ObterOuCriarPainelFonte(aba, m_panelTopologiaId, "TOPOLOGIA");
                RibbonSplitButton dropdownExportarTopologia = CriarDropdown(
                    painelTopologia,
                    "ZAGO_DROPDOWN_TOPOLOGIA_EXPORTAR",
                    "Exportar",
                    "CE",
                    "Comandos de exportação da topologia da rede a partir dos elementos do Civil 3D.");
                AdicionarItemDropdownDummy(
                    dropdownExportarTopologia,
                    "ZAGO_EXPORTAR_CSV_CANAIS",
                    "CSV Canais e Bueiros",
                    "TOPOLOGIA > CSV CANAIS E BUEIROS",
                    "CE",
                    "Exporta CSV com a topologia de canais e bueiros (em definição).");

                // DELETAR
                RibbonPanelSource painelDeletar = ObterOuCriarPainelFonte(aba, m_panelDeletarId, "DELETAR");
                RibbonSplitButton dropdownDeletar = CriarDropdown(
                    painelDeletar,
                    "ZAGO_DROPDOWN_DELETAR",
                    "Deletar",
                    "DL",
                    "Comandos de exclusão de elementos.");
                AdicionarItemDropdownDummy(
                    dropdownDeletar,
                    "ZAGO_DELETAR_DUMMY",
                    "Deletar (Dummy)",
                    "DELETAR > FUNCOES EM DEFINICAO",
                    "DL",
                    "Funções de exclusão (em definição).");

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

        /// <summary>
        /// Cria um dropdown (RibbonSplitButton) dentro de um painel do ribbon.
        /// Configurado como menu puro (IsSplit = false) — o botão inteiro abre
        /// a lista de itens sem executar nenhum comando padrão.
        /// </summary>
        private static RibbonSplitButton CriarDropdown(
            RibbonPanelSource fontePainel,
            string idDropdown,
            string textoDropdown,
            string siglaIcone,
            string descricao = "")
        {
            RibbonSplitButton? dropdownExistente = fontePainel.Items
                .OfType<RibbonSplitButton>()
                .FirstOrDefault(b => string.Equals(b.Id, idDropdown, StringComparison.OrdinalIgnoreCase));

            if (dropdownExistente != null)
                return dropdownExistente;

            var dropdown = new RibbonSplitButton
            {
                Id = idDropdown,
                Text = textoDropdown,
                ShowText = true,
                ShowImage = true,
                Size = RibbonItemSize.Large,
                Orientation = System.Windows.Controls.Orientation.Vertical,
                LargeImage = CriarIcone(siglaIcone, grande: true),
                Image = CriarIcone(siglaIcone, grande: false),
                ListStyle = RibbonSplitButtonListStyle.List,
                IsSplit = false,
                IsSynchronizedWithCurrentItem = false
            };

            if (!string.IsNullOrWhiteSpace(descricao))
            {
                dropdown.Description = descricao;
                dropdown.ToolTip = new RibbonToolTip
                {
                    Title = textoDropdown.Replace("\n", " "),
                    Content = descricao,
                    Command = idDropdown
                };
            }

            fontePainel.Items.Add(dropdown);
            return dropdown;
        }

        private static void AdicionarItemDropdown(
            RibbonSplitButton dropdown,
            string idItem,
            string textoItem,
            string nomeComando,
            string siglaIcone,
            string descricao = "")
        {
            AdicionarItem(dropdown, idItem, textoItem, nomeComando, siglaIcone, descricao);
        }

        private static void AdicionarItemDropdownDummy(
            RibbonSplitButton dropdown,
            string idItem,
            string textoItem,
            string mensagemDummy,
            string siglaIcone,
            string descricao = "")
        {
            AdicionarItem(dropdown, idItem, textoItem, m_prefixoComandoDummy + mensagemDummy, siglaIcone, descricao);
        }

        private static void AdicionarItem(
            RibbonSplitButton dropdown,
            string idItem,
            string textoItem,
            string parametroComando,
            string siglaIcone,
            string descricao = "")
        {
            bool itemExiste = dropdown.Items
                .OfType<RibbonButton>()
                .Any(b => string.Equals(b.Id, idItem, StringComparison.OrdinalIgnoreCase));

            if (itemExiste)
                return;

            var item = new RibbonButton
            {
                Id = idItem,
                Text = textoItem,
                ShowText = true,
                ShowImage = true,
                Size = RibbonItemSize.Large,
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                LargeImage = CriarIcone(siglaIcone, grande: true),
                Image = CriarIcone(siglaIcone, grande: false),
                CommandParameter = parametroComando,
                CommandHandler = new ManipuladorComandoRibbon()
            };

            if (!string.IsNullOrWhiteSpace(descricao))
            {
                item.Description = descricao;
                item.ToolTip = new RibbonToolTip
                {
                    Title = textoItem.Replace("\n", " "),
                    Content = descricao,
                    Command = idItem
                };
            }

            dropdown.Items.Add(item);
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

                case "BC":
                    // Bacia (catchment) a partir de hatch: contorno fechado irregular preenchido,
                    // com uma polilinha interna (talvegue/flow path) saindo de um ponto para o outro.
                    {
                        var contorno = new StreamGeometry();
                        using (var ctx = contorno.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.14, s * 0.30), true, true);
                            ctx.LineTo(new Point(s * 0.40, s * 0.12), true, false);
                            ctx.LineTo(new Point(s * 0.72, s * 0.18), true, false);
                            ctx.LineTo(new Point(s * 0.90, s * 0.42), true, false);
                            ctx.LineTo(new Point(s * 0.82, s * 0.74), true, false);
                            ctx.LineTo(new Point(s * 0.50, s * 0.90), true, false);
                            ctx.LineTo(new Point(s * 0.20, s * 0.76), true, false);
                            ctx.LineTo(new Point(s * 0.10, s * 0.54), true, false);
                        }
                        contorno.Freeze();
                        g.DrawGeometry(corClara, penPrincipal, contorno);

                        // Talvegue (flow path) ziguezagueando de um extremo ao outro da bacia.
                        var talvegue = new StreamGeometry();
                        using (var ctx = talvegue.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.32, s * 0.28), false, false);
                            ctx.LineTo(new Point(s * 0.48, s * 0.46), true, false);
                            ctx.LineTo(new Point(s * 0.40, s * 0.62), true, false);
                            ctx.LineTo(new Point(s * 0.62, s * 0.78), true, false);
                        }
                        talvegue.Freeze();
                        g.DrawGeometry(null, penEscuro, talvegue);

                        // Ponto de descarga (ponta da flow path).
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.62, s * 0.78), s * 0.08, s * 0.08);
                    }
                    break;

                case "TC":
                    // Talvegue dentro do catchment:
                    // - poligono curvo preenchido representa a subbacia (boundary do catchment)
                    // - polilinha em ziguezague descendente marcada por pontos representa o talvegue (FlowPath)
                    {
                        var poligono = new StreamGeometry();
                        using (var ctx = poligono.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.15, s * 0.30), true, true);
                            ctx.LineTo(new Point(s * 0.55, s * 0.12), true, false);
                            ctx.LineTo(new Point(s * 0.88, s * 0.35), true, false);
                            ctx.LineTo(new Point(s * 0.82, s * 0.82), true, false);
                            ctx.LineTo(new Point(s * 0.35, s * 0.90), true, false);
                            ctx.LineTo(new Point(s * 0.10, s * 0.62), true, false);
                        }
                        poligono.Freeze();
                        g.DrawGeometry(corClara, penEscuro, poligono);

                        // Talvegue (FlowPath): linha em ziguezague descendente com pontos nos vertices
                        var talvegue = new StreamGeometry();
                        using (var ctx = talvegue.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.28, s * 0.30), false, false);
                            ctx.LineTo(new Point(s * 0.42, s * 0.48), true, false);
                            ctx.LineTo(new Point(s * 0.58, s * 0.55), true, false);
                            ctx.LineTo(new Point(s * 0.72, s * 0.78), true, false);
                        }
                        talvegue.Freeze();
                        g.DrawGeometry(null, penPrincipal, talvegue);

                        double r = s * 0.07;
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.28, s * 0.30), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.72, s * 0.78), r, r);
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

                case "AP":
                    // Converter alinhamento em polilinha: seta apontando de um
                    // alinhamento (curva suave) para uma polilinha (segmentos
                    // retos + vertices destacados), indicando a transformacao.
                    {
                        // Alinhamento de origem (curva suave) na parte superior
                        var origem = new StreamGeometry();
                        using (var ctx = origem.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.08, s * 0.32), false, false);
                            ctx.BezierTo(
                                new Point(s * 0.28, s * 0.10),
                                new Point(s * 0.60, s * 0.40),
                                new Point(s * 0.92, s * 0.20),
                                true, false);
                        }
                        origem.Freeze();
                        g.DrawGeometry(null, penEscuro, origem);

                        // Seta para baixo (transformacao)
                        var seta = new StreamGeometry();
                        using (var ctx = seta.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.50, s * 0.62), true, true);
                            ctx.LineTo(new Point(s * 0.40, s * 0.50), true, false);
                            ctx.LineTo(new Point(s * 0.46, s * 0.50), true, false);
                            ctx.LineTo(new Point(s * 0.46, s * 0.42), true, false);
                            ctx.LineTo(new Point(s * 0.54, s * 0.42), true, false);
                            ctx.LineTo(new Point(s * 0.54, s * 0.50), true, false);
                            ctx.LineTo(new Point(s * 0.60, s * 0.50), true, false);
                        }
                        seta.Freeze();
                        g.DrawGeometry(corCinza, null, seta);

                        // Polilinha 2D resultante (segmentos retos com vertices)
                        var destino = new StreamGeometry();
                        using (var ctx = destino.Open())
                        {
                            ctx.BeginFigure(new Point(s * 0.10, s * 0.88), false, false);
                            ctx.LineTo(new Point(s * 0.32, s * 0.72), true, false);
                            ctx.LineTo(new Point(s * 0.58, s * 0.82), true, false);
                            ctx.LineTo(new Point(s * 0.90, s * 0.68), true, false);
                        }
                        destino.Freeze();
                        g.DrawGeometry(null, penPrincipal, destino);

                        double r = s * 0.075;
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.10, s * 0.88), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.32, s * 0.72), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.58, s * 0.82), r, r);
                        g.DrawEllipse(corPrincipal, null, new Point(s * 0.90, s * 0.68), r, r);
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
                "AP" => Color.FromRgb(32, 150, 200),
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
                "BC" => Color.FromRgb(30, 144, 180),
                "TC" => Color.FromRgb(30, 110, 150),
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
