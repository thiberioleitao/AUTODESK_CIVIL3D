using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Logica de criacao automatica de perfis de projeto (layout profiles)
/// para todos os alinhamentos do desenho, replicando a rotina Dynamo
/// "1.1 - CRIAR PERFIS DE PROJETO A PARTIR DE ALINHAMENTOS".
/// </summary>
public static class PerfilProjetoService
{
    private const double m_toleranciaEstacao = 1e-6;
    private const int m_maximoIteracoesSuavizacao = 10000;

    /// <summary>
    /// Executa o fluxo completo: para cada alinhamento, amostra a superficie,
    /// calcula o perfil de projeto com declividade minima, suaviza e cria o
    /// layout profile no Civil 3D.
    /// </summary>
    public static CriarPerfisProjetoResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        CriarPerfisProjetoRequest request)
    {
        var resultado = new CriarPerfisProjetoResultado();

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
    /// Processa um unico alinhamento: amostra, calcula, suaviza e cria o perfil.
    /// Cada alinhamento usa sua propria transacao para isolamento de falhas.
    /// </summary>
    private static void ProcessarAlinhamento(
        Database db,
        ObjectId idAlinhamento,
        ObjectId idSuperficie,
        ObjectId idCamada,
        ObjectId idEstiloPerfil,
        ObjectId idConjuntoRotulos,
        CriarPerfisProjetoRequest request,
        CriarPerfisProjetoResultado resultado,
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
            string nomePerfil = nomeAlinhamento + request.SufixoPerfil;

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

            etapaAtual = "leitura da superficie";
            TinSurface superficie = (TinSurface)transacao.GetObject(idSuperficie, OpenMode.ForRead);

            // 1. Gerar estacoes
            etapaAtual = "geracao de estacoes";
            List<double> estacoes = GerarEstacoes(alinhamento, request.EspacamentoEstacas);
            resultado.Logs.Add($"'{nomeAlinhamento}': {estacoes.Count} estacoes geradas (inicio={alinhamento.StartingStation:F2}, fim={alinhamento.EndingStation:F2}).");

            // 2. Amostrar superficie
            etapaAtual = "amostragem da superficie";
            List<(double Estacao, double Elevacao)> amostras =
                AmostrarSuperficie(alinhamento, superficie, estacoes, resultado);

            if (amostras.Count < 2)
            {
                resultado.TotalPerfisIgnorados++;
                resultado.Logs.Add($"'{nomeAlinhamento}': apenas {amostras.Count} ponto(s) amostrado(s) de {estacoes.Count}, ignorado.");
                transacao.Commit();
                return;
            }

            // 3. Calcular perfil de projeto
            etapaAtual = "calculo do perfil de projeto";
            List<(double Estacao, double Elevacao)> perfilProjeto =
                CalcularPerfilProjeto(amostras, request.OffsetAbaixoTerreno, request.DeclivedadeMinima);

            resultado.Logs.Add($"'{nomeAlinhamento}': {perfilProjeto.Count} PVIs apos calculo de declividade.");

            if (perfilProjeto.Count < 2)
            {
                resultado.TotalPerfisIgnorados++;
                resultado.Logs.Add($"'{nomeAlinhamento}': perfil resultou em menos de 2 PVIs, ignorado.");
                transacao.Commit();
                return;
            }

            // 4. Suavizar
            etapaAtual = "suavizacao do perfil";
            List<(double Estacao, double Elevacao)> perfilSuavizado =
                SuavizarPerfil(perfilProjeto, request.LimiteVariacaoDeclividade);

            resultado.Logs.Add($"'{nomeAlinhamento}': {perfilSuavizado.Count} PVIs apos suavizacao.");

            // 5. Criar layout profile
            etapaAtual = "criacao do layout profile";
            ObjectId idPerfil = Profile.CreateByLayout(
                nomePerfil,
                alinhamento.ObjectId,
                idCamada,
                idEstiloPerfil,
                idConjuntoRotulos);

            etapaAtual = "insercao dos PVIs";
            Profile perfil = (Profile)transacao.GetObject(idPerfil, OpenMode.ForWrite);

            for (int iPvi = 0; iPvi < perfilSuavizado.Count; iPvi++)
            {
                var ponto = perfilSuavizado[iPvi];
                perfil.PVIs.AddPVI(ponto.Estacao, ponto.Elevacao);
            }

            transacao.Commit();

            resultado.TotalPerfisCriados++;
            resultado.NomesCriados.Add(nomePerfil);
            ed.WriteMessage($"\n[ZagoCivil3D]   Perfil '{nomePerfil}' criado com {perfilSuavizado.Count} PVIs.");
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add($"'{nomeAlinhamento}' - erro na etapa '{etapaAtual}': {ex.Message}");
        }
    }

    /// <summary>
    /// Gera uma lista de estacoes igualmente espacadas do inicio ao fim do alinhamento,
    /// garantindo que a estacao final esteja sempre incluida.
    /// </summary>
    private static List<double> GerarEstacoes(Alignment alinhamento, double espacamento)
    {
        double inicio = alinhamento.StartingStation;
        double fim = alinhamento.EndingStation;
        var estacoes = new List<double> { inicio };

        double atual = inicio;
        while (true)
        {
            double proxima = atual + espacamento;
            if (proxima > fim - m_toleranciaEstacao)
                break;

            estacoes.Add(proxima);
            atual = proxima;
        }

        // Garante estacao final
        if (estacoes.Count == 0 || Math.Abs(estacoes[^1] - fim) > m_toleranciaEstacao)
            estacoes.Add(fim);

        return estacoes;
    }

    /// <summary>
    /// Obtem a elevacao da superficie em cada estacao do alinhamento.
    /// Estacoes fora dos limites da superficie sao ignoradas com log.
    /// </summary>
    private static List<(double Estacao, double Elevacao)> AmostrarSuperficie(
        Alignment alinhamento,
        TinSurface superficie,
        List<double> estacoes,
        CriarPerfisProjetoResultado resultado)
    {
        var amostras = new List<(double, double)>();

        foreach (double estacao in estacoes)
        {
            try
            {
                double easting = 0;
                double northing = 0;
                alinhamento.PointLocation(estacao, 0.0, ref easting, ref northing);

                double elevacao = superficie.FindElevationAtXY(easting, northing);
                amostras.Add((estacao, elevacao));
            }
            catch
            {
                resultado.Logs.Add($"'{alinhamento.Name}': estacao {estacao:F2} fora da superficie, ignorada.");
            }
        }

        return amostras;
    }

    /// <summary>
    /// Calcula o perfil de projeto com declividade minima.
    /// Replica o Script 3 do Dynamo: processa montante -> jusante,
    /// so adiciona ponto se o terreno esta abaixo do minimo de todas
    /// as elevacoes anteriores do projeto, e impoe declividade minima.
    /// Quando declivedadeMinima == 0, o perfil acompanha o terreno
    /// livremente (sobe e desce) sem restricao de declividade.
    /// </summary>
    private static List<(double Estacao, double Elevacao)> CalcularPerfilProjeto(
        List<(double Estacao, double Elevacao)> amostras,
        double offset,
        double declivedadeMinima)
    {
        bool semRestricaoDeclividade = declivedadeMinima <= 0;

        // Quando nao ha restricao de declividade, o perfil acompanha
        // todas as estacoes do terreno (aplicando apenas o offset).
        if (semRestricaoDeclividade)
        {
            var resultado = new List<(double, double)>();
            foreach (var amostra in amostras)
                resultado.Add((amostra.Estacao, amostra.Elevacao - offset));
            return resultado;
        }

        var estacoesProjeto = new List<double>();
        var cotasProjeto = new List<double>();

        for (int i = 0; i < amostras.Count; i++)
        {
            double estacao = amostras[i].Estacao;
            double cotaTerreno = amostras[i].Elevacao - offset;

            if (i == 0)
            {
                estacoesProjeto.Add(estacao);
                cotasProjeto.Add(cotaTerreno);
                continue;
            }

            double cotaMinimaMontante = cotasProjeto.Min();

            if (cotaTerreno >= cotaMinimaMontante)
                continue;

            double dL = estacao - estacoesProjeto[^1];
            if (dL == 0)
                continue;

            double dZ = cotasProjeto[^1] - cotaTerreno;
            double declividade = dZ / dL;

            estacoesProjeto.Add(estacao);

            if (declividade > declivedadeMinima)
            {
                cotasProjeto.Add(cotaTerreno);
            }
            else
            {
                double cotaProjetada = cotasProjeto[^1] - declivedadeMinima * dL;
                cotasProjeto.Add(cotaProjetada);
            }
        }

        // Garantir estacao final
        double estacaoFinal = amostras[^1].Estacao;
        double cotaTerrenoFinal = amostras[^1].Elevacao - offset;

        if (estacoesProjeto.Count == 0)
        {
            estacoesProjeto.Add(amostras[0].Estacao);
            cotasProjeto.Add(amostras[0].Elevacao - offset);
        }

        if (Math.Abs(estacoesProjeto[^1] - estacaoFinal) > m_toleranciaEstacao)
        {
            double dLFinal = estacaoFinal - estacoesProjeto[^1];

            if (dLFinal != 0)
            {
                double dZFinal = cotasProjeto[^1] - cotaTerrenoFinal;
                double declivedadeFinal = dZFinal / dLFinal;

                estacoesProjeto.Add(estacaoFinal);

                if (declivedadeFinal > declivedadeMinima)
                {
                    cotasProjeto.Add(cotaTerrenoFinal);
                }
                else
                {
                    cotasProjeto.Add(cotasProjeto[^1] - declivedadeMinima * dLFinal);
                }
            }
        }

        var resultadoFinal = new List<(double, double)>();
        for (int i = 0; i < estacoesProjeto.Count; i++)
            resultadoFinal.Add((estacoesProjeto[i], cotasProjeto[i]));

        return resultadoFinal;
    }

    /// <summary>
    /// Suaviza o perfil removendo iterativamente PIs intermediarios
    /// cuja variacao de declividade entre segmentos adjacentes e menor
    /// que o limiar. Replica o Script 4 do Dynamo.
    /// Quando limiteVariacao == 0, nao remove nenhum PI (todas as
    /// variacoes de declividade sao consideradas relevantes).
    /// </summary>
    private static List<(double Estacao, double Elevacao)> SuavizarPerfil(
        List<(double Estacao, double Elevacao)> perfil,
        double limiteVariacao)
    {
        // Quando o limite e zero, nenhuma variacao e admissivel para remocao,
        // portanto todos os PIs sao mantidos.
        if (limiteVariacao <= 0 || perfil.Count < 3)
            return new List<(double, double)>(perfil);

        var estacoes = perfil.Select(p => p.Estacao).ToList();
        var cotas = perfil.Select(p => p.Elevacao).ToList();

        int iteracoes = 0;
        bool mudou = true;

        while (mudou && iteracoes < m_maximoIteracoesSuavizacao)
        {
            mudou = false;
            iteracoes++;

            var novasEstacoes = new List<double> { estacoes[0] };
            var novasCotas = new List<double> { cotas[0] };

            int i = 1;
            while (i < estacoes.Count - 1)
            {
                double s0 = novasEstacoes[^1];
                double z0 = novasCotas[^1];
                double s1 = estacoes[i];
                double z1 = cotas[i];
                double s2 = estacoes[i + 1];
                double z2 = cotas[i + 1];

                double dL1 = s1 - s0;
                double dL2 = s2 - s1;

                if (dL1 == 0 || dL2 == 0)
                {
                    novasEstacoes.Add(s1);
                    novasCotas.Add(z1);
                    i++;
                    continue;
                }

                double sAnt = (z0 - z1) / dL1;
                double sPos = (z1 - z2) / dL2;
                double deltaS = Math.Abs(sPos - sAnt);

                if (deltaS < limiteVariacao)
                {
                    // Remove ponto intermediario
                    mudou = true;
                }
                else
                {
                    novasEstacoes.Add(s1);
                    novasCotas.Add(z1);
                }

                i++;
            }

            // Garante ultimo ponto
            novasEstacoes.Add(estacoes[^1]);
            novasCotas.Add(cotas[^1]);

            estacoes = novasEstacoes;
            cotas = novasCotas;
        }

        var resultado = new List<(double, double)>();
        for (int j = 0; j < estacoes.Count; j++)
            resultado.Add((estacoes[j], cotas[j]));

        return resultado;
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
