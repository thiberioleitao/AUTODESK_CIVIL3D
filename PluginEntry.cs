// PluginEntry.cs
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(HelloCivil3D.Commands.CriarAlinhamentosCommand))]

namespace HelloCivil3D
{
    public class PluginEntry : IExtensionApplication
    {
        public void Initialize()
        {
            Ribbon.RibbonInitializer.InitializeRibbon();
        }

        public void Terminate()
        {
        }
    }
}