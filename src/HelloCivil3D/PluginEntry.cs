using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

[assembly: ExtensionApplication(typeof(HelloCivil3D.PluginEntry))]
[assembly: CommandClass(typeof(HelloCivil3D.Commands.CriarAlinhamentosCommand))]

namespace HelloCivil3D
{
    public class PluginEntry : IExtensionApplication
    {
        public void Initialize()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc?.Editor.WriteMessage("\n[HelloCivil3D] Initialize executado.");

            Ribbon.RibbonInitializer.InitializeRibbon();
        }

        public void Terminate()
        {
        }
    }
}