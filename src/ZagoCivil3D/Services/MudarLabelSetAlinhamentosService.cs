using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Logica de aplicacao de um Alignment Label Set Style a todos os alinhamentos
/// do desenho. Replica a rotina Dynamo "MUDAR LABEL SET DOS ALINHAMENTOS":
/// opcionalmente apaga os label groups existentes e em seguida chama
/// ImportLabelSet com o ObjectId do label set escolhido.
/// </summary>
public static class MudarLabelSetAlinhamentosService
{
    /// <summary>
    /// Executa o fluxo completo: para cada alinhamento, apaga os label groups
    /// (se solicitado) e importa o novo label set.
    /// </summary>
    public static MudarLabelSetAlinhamentosResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        MudarLabelSetAlinhamentosRequest request)
    {
        var resultado = new MudarLabelSetAlinhamentosResultado();

        ObjectId idLabelSet = ObterIdLabelSetPorNome(civilDoc, db, request.NomeLabelSetStyle);
        if (idLabelSet.IsNull)
        {
            resultado.MensagensErro.Add(
                $"Label Set Style '{request.NomeLabelSetStyle}' nao encontrado.");
            return resultado;
        }

        ObjectIdCollection idsAlinhamentos = civilDoc.GetAlignmentIds();
        resultado.TotalAlinhamentos = idsAlinhamentos.Count;

        if (idsAlinhamentos.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhum alinhamento encontrado no desenho.");
            return resultado;
        }

        ed.WriteMessage(
            $"\n[ZagoCivil3D] Aplicando label set '{request.NomeLabelSetStyle}' em {idsAlinhamentos.Count} alinhamento(s)...");

        foreach (ObjectId idAlinhamento in idsAlinhamentos)
        {
            ProcessarAlinhamento(db, idAlinhamento, idLabelSet, request, resultado, ed);
        }

        return resultado;
    }

    /// <summary>
    /// Processa um unico alinhamento: apaga os label groups existentes (se
    /// solicitado) e importa o novo label set. Cada alinhamento roda em sua
    /// propria transacao para que um erro isolado nao aborte o lote.
    /// </summary>
    private static void ProcessarAlinhamento(
        Database db,
        ObjectId idAlinhamento,
        ObjectId idLabelSet,
        MudarLabelSetAlinhamentosRequest request,
        MudarLabelSetAlinhamentosResultado resultado,
        Editor ed)
    {
        string nomeAlinhamento = "?";
        string etapaAtual = "abertura";

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            etapaAtual = "leitura do alinhamento";
            if (transacao.GetObject(idAlinhamento, OpenMode.ForWrite) is not Alignment alinhamento)
            {
                resultado.MensagensErro.Add("Objeto nao e um Alignment valido.");
                transacao.Commit();
                return;
            }

            nomeAlinhamento = alinhamento.Name;
            ed.WriteMessage($"\n[ZagoCivil3D] Alinhamento: {nomeAlinhamento}");

            if (request.ApagarExistentes)
            {
                etapaAtual = "exclusao dos label groups existentes";
                int apagados = ApagarLabelGroups(transacao, alinhamento, resultado, nomeAlinhamento);
                resultado.TotalLabelGroupsApagados += apagados;
                if (apagados > 0)
                    resultado.Logs.Add($"'{nomeAlinhamento}': {apagados} label group(s) apagado(s).");
            }

            etapaAtual = "aplicacao do label set";
            alinhamento.ImportLabelSet(idLabelSet);

            transacao.Commit();

            resultado.TotalAplicadosComSucesso++;
            resultado.NomesProcessados.Add(nomeAlinhamento);
            ed.WriteMessage($"\n[ZagoCivil3D]   Label set aplicado em '{nomeAlinhamento}'.");
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add(
                $"'{nomeAlinhamento}' - erro na etapa '{etapaAtual}': {ex.Message}");
        }
    }

    /// <summary>
    /// Apaga todos os label groups associados a um alinhamento. Retorna a
    /// quantidade efetivamente apagada.
    /// </summary>
    private static int ApagarLabelGroups(
        Transaction transacao,
        Alignment alinhamento,
        MudarLabelSetAlinhamentosResultado resultado,
        string nomeAlinhamento)
    {
        int apagados = 0;

        ObjectIdCollection idsLabelGroups;
        try
        {
            idsLabelGroups = alinhamento.GetAlignmentLabelGroupIds();
        }
        catch (Exception ex)
        {
            resultado.Logs.Add(
                $"'{nomeAlinhamento}': nao foi possivel listar label groups ({ex.Message}).");
            return 0;
        }

        foreach (ObjectId idLabelGroup in idsLabelGroups)
        {
            if (idLabelGroup.IsNull || idLabelGroup.IsErased) continue;

            try
            {
                if (transacao.GetObject(idLabelGroup, OpenMode.ForWrite) is Autodesk.AutoCAD.DatabaseServices.DBObject objLabel)
                {
                    objLabel.Erase();
                    apagados++;
                }
            }
            catch (Exception ex)
            {
                resultado.Logs.Add(
                    $"'{nomeAlinhamento}': label group nao apagado ({ex.Message}).");
            }
        }

        return apagados;
    }

    /// <summary>
    /// Localiza o Alignment Label Set Style pelo nome (case-insensitive).
    /// </summary>
    private static ObjectId ObterIdLabelSetPorNome(CivilDocument civilDoc, Database db, string nome)
    {
        if (string.IsNullOrWhiteSpace(nome))
            return ObjectId.Null;

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is StyleBase estilo
                    && string.Equals(estilo.Name, nome, StringComparison.OrdinalIgnoreCase))
                {
                    transacao.Commit();
                    return id;
                }
            }
            catch { }
        }

        transacao.Commit();
        return ObjectId.Null;
    }
}
