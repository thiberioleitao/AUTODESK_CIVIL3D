using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using ZagoCivil3D.Commands;

[assembly: ExtensionApplication(typeof(ZagoCivil3D.PluginEntry))]
[assembly: CommandClass(typeof(ZagoCivil3D.PluginEntry))]
[assembly: CommandClass(typeof(CriarAlinhamentosCommand))]
[assembly: CommandClass(typeof(CriarAlinhamentosOrdenadosCommand))]
[assembly: CommandClass(typeof(TerraplenagemFeatureLinesCommand))]
[assembly: CommandClass(typeof(CriarPerfisProjetoCommand))]

namespace ZagoCivil3D;

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
        EscreverLog("Initialize iniciado via NETLOAD.");

        //Classe RibbonInitializer dentro do namespace Ribbon
        Ribbon.RibbonInitializer.InicializarRibbon();

        EscreverLog("Initialize finalizado.");
    }

    /// <summary>
    /// Finaliza o plugin.
    /// </summary>
    public void Terminate()
    {
        EscreverLog("Terminate chamado.");
    }

    /// <summary>
    /// Comando auxiliar para forçar a recriação da ribbon durante debug.
    /// </summary>
    [CommandMethod("PL_DEBUG_RIBBON")]
    public void DepurarRibbon()
    {
        EscreverLog("PL_DEBUG_RIBBON executado.");
        Ribbon.RibbonInitializer.InicializarRibbon();
    }

    private static void EscreverLog(string mensagem)
    {
        var documento = Application.DocumentManager.MdiActiveDocument;
        documento?.Editor.WriteMessage($"\n[ZagoCivil3D] {mensagem}");
    }
}
