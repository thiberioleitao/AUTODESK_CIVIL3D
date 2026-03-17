// RibbonInitializer.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using System;
using System.Linq;
using System.Runtime.Versioning;

namespace HelloCivil3D.Ribbon
{
    [SupportedOSPlatform("windows")]
    public static class RibbonInitializer
    {
        private const string TabId = "PL_ENGENHARIA_TAB";
        private const string PanelSourceId = "PL_DRENAGEM_PANEL";

        public static void InitializeRibbon()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage("\n[HelloCivil3D] InitializeRibbon chamado.");

                if (ComponentManager.Ribbon != null)
                {
                    doc?.Editor.WriteMessage("\n[HelloCivil3D] Ribbon já disponível.");
                    CreateRibbon(ComponentManager.Ribbon);
                    return;
                }

                doc?.Editor.WriteMessage("\n[HelloCivil3D] Ribbon ainda não disponível. Aguardando ItemInitialized.");
                ComponentManager.ItemInitialized += OnItemInitialized;
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
                doc?.Editor.WriteMessage("\n[HelloCivil3D] OnItemInitialized disparado.");

                if (ComponentManager.Ribbon == null)
                    return;

                CreateRibbon(ComponentManager.Ribbon);
                ComponentManager.ItemInitialized -= OnItemInitialized;
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
            }
            catch (Exception ex)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage($"\n[HelloCivil3D] Erro em CreateRibbon: {ex.Message}");
            }
        }
    }

    [SupportedOSPlatform("windows")]
    public class RibbonCommandHandler : System.Windows.Input.ICommand
    {
        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => true;

        public void Execute(object? parameter)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc?.Editor.WriteMessage("\n[HelloCivil3D] Botão do ribbon clicado.");

            if (doc == null)
                return;

            doc.SendStringToExecute(parameter?.ToString() ?? "", true, false, true);
        }
    }
}