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
/// Logica de criacao automatica de surface profiles (perfis do terreno natural
/// e de terraplenagem) para todos os alinhamentos do desenho, a partir de uma
/// superficie TIN. Replica a rotina Dynamo 1.1 "CRIAR PERFIS DO TERRENO NATURAL
/// E TERRAPLENAGEM A PARTIR DE ALINHAMENTOS".
///
/// Diferente de PerfilProjetoService (que cria layout profiles com PVIs
/// calculados), este servico usa Profile.CreateFromSurface, que amostra o
/// terreno nativamente pela API do Civil 3D.
/// </summary>
public static class PerfilTerrenoService
{
    /// <summary>
    /// Executa o fluxo completo: para cada alinhamento, cria um surface profile
    /// amostrando a superficie informada.
    /// </summary>
    public static CriarPerfisTerrenoResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        CriarPerfisTerrenoRequest request)
    {
        var resultado = new CriarPerfisTerrenoResultado();

        ObjectId idSuperficie = ObterIdSuperficiePorNome(civilDoc, db, request.NomeSuperficie);
        if (idSuperficie.IsNull)
        {
            resultado.MensagensErro.Add($"Superficie '{request.NomeSuperficie}' nao encontrada.");
            return resultado;
        }

        ObjectId idEstiloPerfil = ObterIdEstiloPerfil(civilDoc, db, request.NomeEstiloPerfil);
        if (idEstiloPerfil.IsNull)
        {
            resultado.MensagensErro.Add($"Estilo de perfil '{request.NomeEstiloPerfil}' nao encontrado.");
            return resultado;
        }

        ObjectId idConjuntoRotulos = ObterIdConjuntoRotulosPerfil(civilDoc, db, request.NomeConjuntoRotulosPerfil);
        if (idConjuntoRotulos.IsNull)
        {
            resultado.MensagensErro.Add($"Conjunto de rotulos '{request.NomeConjuntoRotulosPerfil}' nao encontrado.");
            return resultado;
        }

        ObjectId idCamada = ObterIdCamadaPorNome(db, request.NomeCamada);
        if (idCamada.IsNull)
        {
            resultado.MensagensErro.Add($"Layer '{request.NomeCamada}' nao encontrada.");
            return resultado;
        }

        ObjectIdCollection idsAlinhamentos = civilDoc.GetAlignmentIds();
        resultado.TotalAlinhamentos = idsAlinhamentos.Count;

        if (idsAlinhamentos.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhum alinhamento encontrado no desenho.");
            return resultado;
        }

        ed.WriteMessage($"\n[ZagoCivil3D] Processando {idsAlinhamentos.Count} alinhamento(s)...");

        foreach (ObjectId idAlinhamento in idsAlinhamentos)
        {
            ProcessarAlinhamento(
                db, idAlinhamento, idSuperficie, idCamada,
                idEstiloPerfil, idConjuntoRotulos, request, resultado, ed);
        }

        return resultado;
    }

    /// <summary>
    /// Processa um unico alinhamento, criando o surface profile associado.
    /// Cada alinhamento usa sua propria transacao para isolamento de falhas.
    /// </summary>
    private static void ProcessarAlinhamento(
        Database db,
        ObjectId idAlinhamento,
        ObjectId idSuperficie,
        ObjectId idCamada,
        ObjectId idEstiloPerfil,
        ObjectId idConjuntoRotulos,
        CriarPerfisTerrenoRequest request,
        CriarPerfisTerrenoResultado resultado,
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
            string nomePerfil = request.NomeSuperficie + "-" + nomeAlinhamento;

            ed.WriteMessage($"\n[ZagoCivil3D] Alinhamento: {nomeAlinhamento}");

            // Verificar se perfil ja existe
            etapaAtual = "verificacao de perfil existente";
            ObjectId idPerfilExistente = ObterIdPerfilPorNome(transacao, alinhamento, nomePerfil);

            if (!idPerfilExistente.IsNull)
            {
                if (!request.SubstituirPerfisExistentes)
                {
                    resultado.TotalPerfisIgnorados++;
                    resultado.Logs.Add($"'{nomeAlinhamento}': perfil '{nomePerfil}' ja existe, ignorado.");
                    transacao.Commit();
                    return;
                }

                // Apagar perfil existente antes de criar o novo
                etapaAtual = "exclusao do perfil existente";
                Profile perfilAntigo = (Profile)transacao.GetObject(idPerfilExistente, OpenMode.ForWrite);
                perfilAntigo.Erase();
                resultado.TotalPerfisSubstituidos++;
                resultado.Logs.Add($"'{nomeAlinhamento}': perfil '{nomePerfil}' existente removido.");
            }

            // Criar surface profile via API nativa do Civil 3D
            etapaAtual = "criacao do surface profile";
            ObjectId idPerfil = Profile.CreateFromSurface(
                nomePerfil,
                alinhamento.ObjectId,
                idSuperficie,
                idCamada,
                idEstiloPerfil,
                idConjuntoRotulos);

            transacao.Commit();

            resultado.TotalPerfisCriados++;
            resultado.NomesCriados.Add(nomePerfil);
            ed.WriteMessage($"\n[ZagoCivil3D]   Perfil '{nomePerfil}' criado.");
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add($"'{nomeAlinhamento}' - erro na etapa '{etapaAtual}': {ex.Message}");
        }
    }

    /// <summary>
    /// Retorna o ObjectId do perfil com o nome especificado, ou ObjectId.Null se nao existir.
    /// </summary>
    private static ObjectId ObterIdPerfilPorNome(
        Transaction transacao,
        Alignment alinhamento,
        string nomePerfil)
    {
        foreach (ObjectId idPerfil in alinhamento.GetProfileIds())
        {
            try
            {
                if (transacao.GetObject(idPerfil, OpenMode.ForRead) is Profile perfil
                    && string.Equals(perfil.Name, nomePerfil, StringComparison.OrdinalIgnoreCase))
                {
                    return idPerfil;
                }
            }
            catch
            {
                // Perfil invalido/apagado, ignorar
            }
        }

        return ObjectId.Null;
    }

    /// <summary>
    /// Localiza a superficie TIN pelo nome.
    /// </summary>
    private static ObjectId ObterIdSuperficiePorNome(CivilDocument civilDoc, Database db, string nome)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.GetSurfaceIds())
        {
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is TinSurface superficie
                    && string.Equals(superficie.Name, nome, StringComparison.OrdinalIgnoreCase))
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

    /// <summary>
    /// Localiza o estilo de perfil pelo nome.
    /// </summary>
    private static ObjectId ObterIdEstiloPerfil(CivilDocument civilDoc, Database db, string nome)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.ProfileStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is ProfileStyle estilo
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

    /// <summary>
    /// Localiza o conjunto de rotulos de perfil pelo nome.
    /// </summary>
    private static ObjectId ObterIdConjuntoRotulosPerfil(CivilDocument civilDoc, Database db, string nome)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId id in civilDoc.Styles.LabelSetStyles.ProfileLabelSetStyles)
        {
            if (id.IsNull || id.IsErased) continue;
            try
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is StyleBase conjunto
                    && string.Equals(conjunto.Name, nome, StringComparison.OrdinalIgnoreCase))
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

    /// <summary>
    /// Localiza uma layer pelo nome.
    /// </summary>
    private static ObjectId ObterIdCamadaPorNome(Database db, string nome)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();
        var tabelaCamadas = (LayerTable)transacao.GetObject(db.LayerTableId, OpenMode.ForRead);

        foreach (ObjectId idCamada in tabelaCamadas)
        {
            try
            {
                if (transacao.GetObject(idCamada, OpenMode.ForRead) is LayerTableRecord camada
                    && string.Equals(camada.Name, nome, StringComparison.OrdinalIgnoreCase))
                {
                    transacao.Commit();
                    return idCamada;
                }
            }
            catch { }
        }

        transacao.Commit();
        return ObjectId.Null;
    }
}
