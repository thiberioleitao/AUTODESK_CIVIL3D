using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Logica de criacao de corredores vazios a partir de todos os alinhamentos do
/// desenho. Replica a rotina Dynamo "CRIAR CORREDORES A PARTIR DE ALINHAMENTOS":
/// para cada alinhamento, cria um corredor vazio cujo nome e derivado do nome
/// do alinhamento (com prefixo/sufixo opcionais).
/// </summary>
public static class CorredorCreationService
{
    /// <summary>
    /// Executa o fluxo completo: le a lista de alinhamentos e cria um corredor
    /// vazio para cada um. Cada corredor e criado em sua propria transacao para
    /// que uma falha isolada nao aborte o lote inteiro.
    /// </summary>
    public static CriarCorredoresAlinhamentosResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        CriarCorredoresAlinhamentosRequest request)
    {
        var resultado = new CriarCorredoresAlinhamentosResultado();

        ObjectIdCollection idsAlinhamentos = civilDoc.GetAlignmentIds();
        resultado.TotalAlinhamentos = idsAlinhamentos.Count;

        if (idsAlinhamentos.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhum alinhamento encontrado no desenho.");
            return resultado;
        }

        HashSet<string> nomesCorredoresExistentes = ObterNomesCorredoresExistentes(civilDoc, db);

        ed.WriteMessage(
            $"\n[ZagoCivil3D] Criando corredores para {idsAlinhamentos.Count} alinhamento(s)...");

        string prefixo = request.PrefixoNome ?? string.Empty;
        string sufixo = request.SufixoNome ?? string.Empty;

        foreach (ObjectId idAlinhamento in idsAlinhamentos)
        {
            ProcessarAlinhamento(
                db,
                idAlinhamento,
                prefixo,
                sufixo,
                nomesCorredoresExistentes,
                request.IgnorarExistentes,
                resultado,
                ed);
        }

        return resultado;
    }

    /// <summary>
    /// Processa um unico alinhamento: le o nome, monta o nome do corredor e
    /// cria o corredor vazio. Atualiza o conjunto de nomes existentes para
    /// evitar colisoes dentro do proprio lote.
    /// </summary>
    private static void ProcessarAlinhamento(
        Database db,
        ObjectId idAlinhamento,
        string prefixo,
        string sufixo,
        HashSet<string> nomesExistentes,
        bool ignorarExistentes,
        CriarCorredoresAlinhamentosResultado resultado,
        Editor ed)
    {
        string nomeAlinhamento = "?";
        string etapaAtual = "abertura";

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            etapaAtual = "leitura do alinhamento";
            if (transacao.GetObject(idAlinhamento, OpenMode.ForRead) is not Alignment alinhamento)
            {
                resultado.MensagensErro.Add("Objeto nao e um Alignment valido.");
                transacao.Commit();
                return;
            }

            nomeAlinhamento = alinhamento.Name;
            string nomeCorredor = prefixo + nomeAlinhamento + sufixo;

            ed.WriteMessage($"\n[ZagoCivil3D] Alinhamento: {nomeAlinhamento}");

            etapaAtual = "verificacao de colisao";
            if (nomesExistentes.Contains(nomeCorredor))
            {
                if (ignorarExistentes)
                {
                    resultado.TotalIgnorados++;
                    resultado.Logs.Add(
                        $"'{nomeAlinhamento}': corredor '{nomeCorredor}' ja existe, ignorado.");
                    ed.WriteMessage($"\n[ZagoCivil3D]   Ignorado (ja existe): {nomeCorredor}");
                    transacao.Commit();
                    return;
                }

                resultado.MensagensErro.Add(
                    $"'{nomeAlinhamento}': corredor '{nomeCorredor}' ja existe.");
                transacao.Commit();
                return;
            }

            etapaAtual = "criacao do corredor";
            _ = civilDoc_CorridorCollection_Add(nomeCorredor);

            transacao.Commit();

            nomesExistentes.Add(nomeCorredor);
            resultado.TotalCorredoresCriados++;
            resultado.NomesCriados.Add(nomeCorredor);
            ed.WriteMessage($"\n[ZagoCivil3D]   Corredor criado: {nomeCorredor}");
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add(
                $"'{nomeAlinhamento}' - erro na etapa '{etapaAtual}': {ex.Message}");
        }
    }

    /// <summary>
    /// Cria um corredor vazio atraves da CorridorCollection do documento Civil
    /// ativo. Isolado em um metodo para que a chamada real possa ser ajustada
    /// em um unico ponto caso a assinatura da API mude entre versoes.
    /// </summary>
    private static ObjectId civilDoc_CorridorCollection_Add(string nomeCorredor)
    {
        CivilDocument civilDoc = CivilApplication.ActiveDocument;
        return civilDoc.CorridorCollection.Add(nomeCorredor);
    }

    /// <summary>
    /// Le os nomes dos corredores ja existentes no desenho para permitir
    /// deteccao de colisao antes da criacao. Usa comparacao case-insensitive.
    /// </summary>
    private static HashSet<string> ObterNomesCorredoresExistentes(CivilDocument civilDoc, Database db)
    {
        var nomes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.CorridorCollection)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is Corridor corredor
                    && !string.IsNullOrWhiteSpace(corredor.Name))
                {
                    nomes.Add(corredor.Name);
                }
            }
            catch { }
        }

        transacao.Commit();
        return nomes;
    }
}
