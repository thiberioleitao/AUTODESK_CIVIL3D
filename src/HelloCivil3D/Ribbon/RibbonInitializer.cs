// RibbonInitializer.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using System;
using System.Linq;
using System.Runtime.Versioning;

namespace HelloCivil3D.Ribbon
{
    [SupportedOSPlatform("windows")]
    /// <summary>
    /// Centraliza a criação da aba personalizada no Ribbon do Civil 3D.
    /// Mantém lógica idempotente para evitar abas duplicadas em recargas.
    /// </summary>
    public static class RibbonInitializer
    {
        private const string TabId = "PL_ENGENHARIA_TAB";
        private const string PanelSourceId = "PL_DRENAGEM_PANEL";
        private static bool _aguardandoRibbon;

        public static void InitializeRibbon()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage("\n[HelloCivil3D] InitializeRibbon chamado.");
                doc?.Editor.WriteMessage($"\n[HelloCivil3D] Ribbon null? {ComponentManager.Ribbon == null}");

                if (ComponentManager.Ribbon != null)
                {
                    doc?.Editor.WriteMessage("\n[HelloCivil3D] Ribbon já disponível.");
                    CreateRibbon(ComponentManager.Ribbon);
                    return;
                }

                // O Ribbon pode ainda não existir durante o NETLOAD.
                // Nesse caso aguardamos o evento de inicialização da UI.
                if (_aguardandoRibbon)
                {
                    doc?.Editor.WriteMessage("\n[HelloCivil3D] Já aguardando ItemInitialized. Nenhuma nova inscrição realizada.");
                    return;
                }

                doc?.Editor.WriteMessage("\n[HelloCivil3D] Ribbon ainda não disponível. Inscrevendo ItemInitialized.");
                ComponentManager.ItemInitialized += OnItemInitialized;
                _aguardandoRibbon = true;
            }
            catch (Exception ex)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage($"\n[HelloCivil3D] Erro em InitializeRibbon: {ex.Message}");
            }
        }

        private static void OnItemInitialized(object? sender, RibbonItemEventArgs e)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage($"\n[HelloCivil3D] OnItemInitialized disparado. Item: {e.Item?.Id ?? "(sem id)"}");

                if (ComponentManager.Ribbon == null)
                {
                    doc?.Editor.WriteMessage("\n[HelloCivil3D] OnItemInitialized: Ribbon ainda nula.");
                    return;
                }

                CreateRibbon(ComponentManager.Ribbon);
                ComponentManager.ItemInitialized -= OnItemInitialized;
                _aguardandoRibbon = false;
                doc?.Editor.WriteMessage("\n[HelloCivil3D] ItemInitialized removido.");
            }
            catch (Exception ex)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage($"\n[HelloCivil3D] Erro em OnItemInitialized: {ex.Message}");
            }
        }

        private static void CreateRibbon(RibbonControl ribbon)
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage("\n[HelloCivil3D] CreateRibbon chamado.");
                doc?.Editor.WriteMessage($"\n[HelloCivil3D] Total de abas atuais: {ribbon.Tabs.Count}");

                var existingTab = ribbon.Tabs.FirstOrDefault(t => t.Id == TabId);
                if (existingTab != null)
                {
                    doc?.Editor.WriteMessage("\n[HelloCivil3D] Aba já existe.");
                    return;
                }

                var tab = new RibbonTab
                {
                    Title = "PL Engenharia",
                    Id = TabId
                };

                var panelSource = new RibbonPanelSource
                {
                    Title = "Drenagem",
                    Id = PanelSourceId
                };

                var panel = new RibbonPanel
                {
                    Source = panelSource
                };

                var button = new RibbonButton
                {
                    Text = "Criar\nAlinhamentos",
                    ShowText = true,
                    Size = RibbonItemSize.Large,
                    Orientation = System.Windows.Controls.Orientation.Vertical,
                    CommandParameter = "PL_CRIAR_ALINHAMENTOS ",
                    CommandHandler = new RibbonCommandHandler()
                };

                panelSource.Items.Add(button);
                tab.Panels.Add(panel);
                ribbon.Tabs.Add(tab);
                ribbon.ActiveTab = tab;

                doc?.Editor.WriteMessage("\n[HelloCivil3D] Aba e botão criados com sucesso.");
                doc?.Editor.WriteMessage($"\n[HelloCivil3D] Aba ativa: {ribbon.ActiveTab?.Id ?? "(nenhuma)"}");
            }
            catch (Exception ex)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage($"\n[HelloCivil3D] Erro em CreateRibbon: {ex.Message}");
            }
        }
    }

    [SupportedOSPlatform("windows")]
    /// <summary>
    /// Encaminha o clique do botão para um comando de texto do AutoCAD.
    /// </summary>
    public class RibbonCommandHandler : System.Windows.Input.ICommand
    {
        private const string CommandName = "PL_CRIAR_ALINHAMENTOS ";

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => true;

        public void Execute(object? parameter)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc?.Editor.WriteMessage("\n[HelloCivil3D] Botão do ribbon clicado.");

            if (doc == null)
                return;

            string commandText = CommandName;

            string parameterType = parameter?.GetType().FullName ?? "(null)";
            doc?.Editor.WriteMessage($"\n[HelloCivil3D] Tipo de parâmetro no clique: {parameterType}");

            doc?.Editor.WriteMessage($"\n[HelloCivil3D] Enviando comando: {commandText.Trim()}");
            doc.SendStringToExecute(commandText, true, false, true);
        }
    }
}