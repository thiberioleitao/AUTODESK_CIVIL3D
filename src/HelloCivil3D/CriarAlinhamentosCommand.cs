using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

namespace HelloCivil3D.Commands
{
    public class CriarAlinhamentosCommand
    {
        [CommandMethod("PL_CRIAR_ALINHAMENTOS")]
        public void Executar()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
                return;

            doc.Editor.WriteMessage("\n[HelloCivil3D] Comando executado com sucesso.");
        }
    }
}