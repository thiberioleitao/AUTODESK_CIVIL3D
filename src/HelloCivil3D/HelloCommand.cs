using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace HelloCivil3D;

/// <summary>
/// Contem o comando de exemplo que pode ser carregado no Autodesk Civil 3D.
/// </summary>
public class HelloCommand
{
    /// <summary>
    /// Registra o comando "HELLO_CIVIL3D" no AutoCAD/Civil 3D.
    /// Quando o usuario executa esse comando, a mensagem de exemplo e escrita
    /// na linha de comando da aplicacao.
    /// </summary>
    [CommandMethod("HELLO_CIVIL3D")]
    public void HelloCivil3D()
    {
        // Representa o documento ativo no Civil 3D/AutoCAD.
        Document activeDocument = Application.DocumentManager.MdiActiveDocument;

        // Representa o banco de dados do desenho atual.
        // Ele nao e obrigatorio para a mensagem simples, mas foi incluido
        // para demonstrar o uso da namespace Autodesk.AutoCAD.DatabaseServices.
        Database currentDatabase = activeDocument.Database;

        // Representa o editor responsavel pela Command Line do AutoCAD.
        Editor editor = activeDocument.Editor;

        // Evita alerta do compilador sobre a variavel local nao utilizada.
        _ = currentDatabase;

        // Escreve a mensagem solicitada na linha de comando do Civil 3D.
        editor.WriteMessage("\nHello Civil 3D from plugin");
    }
}
