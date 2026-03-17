using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using HelloCivil3D.Commands;

[assembly: ExtensionApplication(typeof(HelloCivil3D.PluginEntry))]
[assembly: CommandClass(typeof(HelloCivil3D.PluginEntry))]
[assembly: CommandClass(typeof(CriarAlinhamentosCommand))]

namespace HelloCivil3D;

/// <summary>
/// Ponto de entrada do plugin para inicialização no NETLOAD.
/// Esse tipo existe para garantir que a aba seja criada mesmo sem executar
/// nenhum comando manual logo após o carregamento da DLL.
/// </summary>
public sealed class PluginEntry : IExtensionApplication
{
    /// <summary>
    /// Inicializa o plugin e dispara a criação da ribbon.
    /// </summary>
    public void Initialize()
    {
        WriteLog("Initialize iniciado via NETLOAD.");

        //Classe RibbonInitializer dentro do namespace Ribbon
        Ribbon.RibbonInitializer.InitializeRibbon();

        WriteLog("Initialize finalizado.");
    }

    /// <summary>
    /// Finaliza o plugin.
    /// </summary>
    public void Terminate()
    {
        WriteLog("Terminate chamado.");
    }

    /// <summary>
    /// Comando auxiliar para forçar a recriação da ribbon durante debug.
    /// </summary>
    [CommandMethod("PL_DEBUG_RIBBON")]
    public void DebugRibbon()
    {
        WriteLog("PL_DEBUG_RIBBON executado.");
        Ribbon.RibbonInitializer.InitializeRibbon();
    }

    private static void WriteLog(string message)
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        doc?.Editor.WriteMessage($"\n[HelloCivil3D] {message}");
    }
}
