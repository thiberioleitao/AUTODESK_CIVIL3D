using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Logica de criacao de CogoPoints nos cruzamentos entre alinhamentos.
/// Porta da rotina Dynamo "CRIAR PONTOS NOS CRUZAMENTOS ENTRE ALINHAMENTOS":
/// coleta pares de alinhamentos, calcula as interseccoes no plano XY, remove
/// duplicados, agrupa em blocos (trackers) com rotulos sequenciais e cria
/// os CogoPoints com RawDescription.
/// </summary>
public static class CriarPontosCruzamentosService
{
    private const int m_limiteAvisosDebug = 10;

    /// <summary>
    /// Executa o fluxo completo.
    /// </summary>
    public static CriarPontosCruzamentosResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        CriarPontosCruzamentosRequest request)
    {
        var resultado = new CriarPontosCruzamentosResultado();

        ObjectIdCollection idsAlinhamentos = civilDoc.GetAlignmentIds();
        resultado.TotalAlinhamentos = idsAlinhamentos.Count;

        if (idsAlinhamentos.Count < 2)
        {
            resultado.MensagensErro.Add(
                "Sao necessarios ao menos dois alinhamentos no desenho para calcular cruzamentos.");
            return resultado;
        }

        ed.WriteMessage(
            $"\n[ZagoCivil3D] Calculando cruzamentos entre {idsAlinhamentos.Count} alinhamento(s)...");

        List<Point3d> pontosBrutos = CalcularPontosBrutosCruzamento(
            db, idsAlinhamentos, resultado);

        resultado.TotalPontosBrutos = pontosBrutos.Count;

        List<Point3d> pontosUnicos = RemoverDuplicadosXy(pontosBrutos, request.ToleranciaXyDuplicados);
        resultado.TotalPontosUnicos = pontosUnicos.Count;

        int duplicadosRemovidos = pontosBrutos.Count - pontosUnicos.Count;
        resultado.Logs.Add(
            $"Pontos brutos: {pontosBrutos.Count}, unicos: {pontosUnicos.Count}, duplicados removidos: {duplicadosRemovidos}.");

        if (pontosUnicos.Count == 0)
        {
            resultado.MensagensErro.Add("Nenhum cruzamento encontrado entre os alinhamentos.");
            return resultado;
        }

        List<Point3d> pontosOrdenados;
        List<string> rotulosOrdenados;
        OrganizarEmBlocos(
            pontosUnicos,
            request,
            resultado,
            out pontosOrdenados,
            out rotulosOrdenados);

        resultado.TotalPontosOrdenados = pontosOrdenados.Count;

        if (!request.CriarCogoPoints)
        {
            resultado.Logs.Add("Criacao de CogoPoints desabilitada (modo dry-run).");
            return resultado;
        }

        CriarCogoPoints(db, civilDoc, pontosUnicos, pontosOrdenados, rotulosOrdenados, request, resultado);

        return resultado;
    }

    /// <summary>
    /// Calcula os pontos de cruzamento (no plano XY) entre todos os pares unicos
    /// de alinhamentos. Cada par e avaliado em uma transacao propria para isolar
    /// falhas. Os pontos retornados ainda podem conter duplicados entre pares.
    /// </summary>
    private static List<Point3d> CalcularPontosBrutosCruzamento(
        Database db,
        ObjectIdCollection idsAlinhamentos,
        CriarPontosCruzamentosResultado resultado)
    {
        var pontos = new List<Point3d>();
        var planoXy = new Plane(Point3d.Origin, Vector3d.ZAxis);
        var ids = idsAlinhamentos.Cast<ObjectId>().ToList();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            int totalPares = 0;
            int paresComIntersecao = 0;

            for (int i = 0; i < ids.Count; i++)
            {
                if (transacao.GetObject(ids[i], OpenMode.ForRead) is not Alignment alinhamento1)
                    continue;

                for (int j = i + 1; j < ids.Count; j++)
                {
                    if (transacao.GetObject(ids[j], OpenMode.ForRead) is not Alignment alinhamento2)
                        continue;

                    totalPares++;

                    var pontosPar = new Point3dCollection();
                    try
                    {
                        alinhamento1.IntersectWith(
                            alinhamento2,
                            Intersect.OnBothOperands,
                            planoXy,
                            pontosPar,
                            IntPtr.Zero,
                            IntPtr.Zero);
                    }
                    catch (Exception ex)
                    {
                        if (resultado.Avisos.Count < m_limiteAvisosDebug)
                        {
                            resultado.Avisos.Add(
                                $"Erro ao intersectar '{alinhamento1.Name}' x '{alinhamento2.Name}': {ex.Message}");
                        }
                        continue;
                    }

                    if (pontosPar.Count == 0)
                        continue;

                    paresComIntersecao++;

                    foreach (Point3d p in pontosPar)
                        pontos.Add(p);
                }
            }

            resultado.TotalParesTestados = totalPares;
            resultado.TotalParesComIntersecao = paresComIntersecao;
        }
        finally
        {
            transacao.Commit();
        }

        return pontos;
    }

    /// <summary>
    /// Remove duplicados no plano XY usando uma grade de tolerancia. Pontos
    /// com distancia XY menor ou igual a <paramref name="tolerancia"/> em
    /// relacao a um ponto ja aceito sao descartados.
    /// </summary>
    private static List<Point3d> RemoverDuplicadosXy(IReadOnlyList<Point3d> pontos, double tolerancia)
    {
        if (tolerancia <= 0)
            tolerancia = 0.01;

        var unicos = new List<Point3d>();
        var grade = new Dictionary<(int, int), List<Point3d>>();
        double toleranciaQuadrado = tolerancia * tolerancia;

        foreach (Point3d ponto in pontos)
        {
            (int, int) chave = (
                (int)Math.Round(ponto.X / tolerancia),
                (int)Math.Round(ponto.Y / tolerancia));

            if (!grade.TryGetValue(chave, out List<Point3d>? lista))
            {
                lista = new List<Point3d>();
                grade[chave] = lista;
                lista.Add(ponto);
                unicos.Add(ponto);
                continue;
            }

            bool duplicado = false;
            foreach (Point3d existente in lista)
            {
                double dx = existente.X - ponto.X;
                double dy = existente.Y - ponto.Y;
                if (dx * dx + dy * dy <= toleranciaQuadrado)
                {
                    duplicado = true;
                    break;
                }
            }

            if (!duplicado)
            {
                lista.Add(ponto);
                unicos.Add(ponto);
            }
        }

        return unicos;
    }

    /// <summary>
    /// Organiza os pontos em blocos sequenciais baseados no comprimento do
    /// tracker, seguindo a mesma logica da rotina Dynamo original:
    /// 1. Escolhe o ponto base (mais a norte e mais a oeste, ou alternando
    ///    para nordeste quando <see cref="CriarPontosCruzamentosRequest.AlternarLadoInicial"/>
    ///    estiver ativo).
    /// 2. Coleta todos os pontos com Y entre <c>Ybase - L</c> e <c>Ybase</c>.
    /// 3. Agrupa os candidatos em colunas por proximidade de X.
    /// 4. Ordena cada coluna de norte para sul.
    /// 5. Atribui rotulos sequenciais e remove os pontos ja tratados do
    ///    conjunto pendente.
    /// </summary>
    private static void OrganizarEmBlocos(
        IReadOnlyList<Point3d> pontosUnicos,
        CriarPontosCruzamentosRequest request,
        CriarPontosCruzamentosResultado resultado,
        out List<Point3d> pontosOrdenados,
        out List<string> rotulosOrdenados)
    {
        pontosOrdenados = new List<Point3d>();
        rotulosOrdenados = new List<string>();

        var pendentes = new List<Point3d>(pontosUnicos);
        int contadorGlobal = 1;
        int indiceBloco = 1;

        while (pendentes.Count > 0)
        {
            (Point3d pontoBase, string ladoInicial) = ObterPontoBase(
                pendentes, indiceBloco, request.AlternarLadoInicial);

            double limiteInferiorY = pontoBase.Y - request.ComprimentoTracker;
            var candidatos = pendentes
                .Where(p => p.Y >= limiteInferiorY && p.Y <= pontoBase.Y + 1e-9)
                .ToList();

            if (candidatos.Count == 0)
            {
                resultado.Avisos.Add($"Bloco {indiceBloco} sem pontos candidatos.");
                break;
            }

            List<List<Point3d>> colunas = AgruparPorColuna(candidatos, request.ToleranciaXColuna);

            if (ladoInicial == "leste")
                colunas.Reverse();

            var bloco = new BlocoCruzamento
            {
                IndiceBloco = indiceBloco,
                LadoInicial = ladoInicial,
                PontoBaseX = pontoBase.X,
                PontoBaseY = pontoBase.Y
            };

            for (int indiceColuna = 0; indiceColuna < colunas.Count; indiceColuna++)
            {
                List<Point3d> colunaOrdenada = colunas[indiceColuna]
                    .OrderByDescending(p => p.Y)
                    .ToList();

                if (colunaOrdenada.Count != request.PontosPorTracker)
                {
                    resultado.Avisos.Add(
                        $"Bloco {indiceBloco} coluna {indiceColuna + 1}: {colunaOrdenada.Count} ponto(s) (esperado {request.PontosPorTracker}).");
                }

                foreach (Point3d ponto in colunaOrdenada)
                {
                    string rotulo = FormatarRotulo(contadorGlobal, request.Prefixo);
                    pontosOrdenados.Add(ponto);
                    rotulosOrdenados.Add(rotulo);
                    bloco.Rotulos.Add(rotulo);
                    contadorGlobal++;
                }
            }

            bloco.QuantidadePontos = bloco.Rotulos.Count;
            resultado.Blocos.Add(bloco);

            var chavesTratadas = new HashSet<(double, double)>(
                colunas.SelectMany(c => c).Select(ChaveFinaXy));

            pendentes = pendentes
                .Where(p => !chavesTratadas.Contains(ChaveFinaXy(p)))
                .ToList();

            indiceBloco++;
        }
    }

    /// <summary>
    /// Escolhe o ponto base do bloco: mais a norte e, entre os mais ao norte,
    /// o mais a oeste (ou mais a leste, quando o bloco deve comecar pela
    /// direita). Retorna tambem o lado de partida do bloco.
    /// </summary>
    private static (Point3d pontoBase, string ladoInicial) ObterPontoBase(
        IReadOnlyList<Point3d> pendentes,
        int indiceBloco,
        bool alternarLadoInicial)
    {
        bool comecaLeste = alternarLadoInicial && (indiceBloco % 2 == 0);

        if (comecaLeste)
        {
            Point3d pontoLeste = pendentes
                .OrderByDescending(p => p.Y)
                .ThenByDescending(p => p.X)
                .First();
            return (pontoLeste, "leste");
        }

        Point3d pontoOeste = pendentes
            .OrderByDescending(p => p.Y)
            .ThenBy(p => p.X)
            .First();
        return (pontoOeste, "oeste");
    }

    /// <summary>
    /// Agrupa pontos em colunas por proximidade de X. Pontos sao ordenados
    /// por X crescente e um novo grupo e iniciado quando a diferenca em
    /// relacao ao X de referencia excede <paramref name="toleranciaX"/>.
    /// </summary>
    private static List<List<Point3d>> AgruparPorColuna(IEnumerable<Point3d> pontos, double toleranciaX)
    {
        var ordenados = pontos.OrderBy(p => p.X).ToList();
        var colunas = new List<List<Point3d>>();

        if (ordenados.Count == 0)
            return colunas;

        var colunaAtual = new List<Point3d>();
        double xReferencia = 0;

        foreach (Point3d ponto in ordenados)
        {
            if (colunaAtual.Count == 0)
            {
                colunaAtual.Add(ponto);
                xReferencia = ponto.X;
                continue;
            }

            if (Math.Abs(ponto.X - xReferencia) <= toleranciaX)
            {
                colunaAtual.Add(ponto);
            }
            else
            {
                colunas.Add(colunaAtual);
                colunaAtual = new List<Point3d> { ponto };
                xReferencia = ponto.X;
            }
        }

        if (colunaAtual.Count > 0)
            colunas.Add(colunaAtual);

        return colunas;
    }

    /// <summary>
    /// Chave fina de identificacao de um ponto pelo par XY arredondado a 8
    /// casas, usada para cruzar o ponto ordenado com o ponto bruto na fase
    /// de aplicacao da RawDescription.
    /// </summary>
    private static (double, double) ChaveFinaXy(Point3d ponto)
    {
        return (Math.Round(ponto.X, 8), Math.Round(ponto.Y, 8));
    }

    /// <summary>Formata o rotulo sequencial no estilo "P01", "P02"...</summary>
    private static string FormatarRotulo(int numero, string prefixo)
    {
        return string.Format(CultureInfo.InvariantCulture, "{0}{1:00}", prefixo ?? string.Empty, numero);
    }

    /// <summary>
    /// Cria os CogoPoints e aplica layer + RawDescription ordenada. A criacao
    /// dos pontos e a aplicacao da descricao ficam em uma unica transacao
    /// para permitir rollback caso algo falhe no meio do processamento.
    /// </summary>
    private static void CriarCogoPoints(
        Database db,
        CivilDocument civilDoc,
        IReadOnlyList<Point3d> pontosUnicos,
        IReadOnlyList<Point3d> pontosOrdenados,
        IReadOnlyList<string> rotulosOrdenados,
        CriarPontosCruzamentosRequest request,
        CriarPontosCruzamentosResultado resultado)
    {
        GarantirLayerExiste(db, request.NomeLayer, resultado);

        var mapaIdPorChave = new Dictionary<(double, double), ObjectId>();

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            CogoPointCollection cogoPoints = civilDoc.CogoPoints;

            for (int indice = 0; indice < pontosUnicos.Count; indice++)
            {
                Point3d pontoAcad = pontosUnicos[indice];

                ObjectId idPonto;
                try
                {
                    idPonto = cogoPoints.Add(pontoAcad, request.Descricao, true);
                }
                catch (Exception ex)
                {
                    resultado.MensagensErro.Add(
                        $"Erro ao criar CogoPoint em ({pontoAcad.X:F3}, {pontoAcad.Y:F3}): {ex.Message}");
                    continue;
                }

                resultado.TotalCogoPointsCriados++;
                mapaIdPorChave[ChaveFinaXy(pontoAcad)] = idPonto;

                if (!string.IsNullOrWhiteSpace(request.NomeLayer))
                {
                    try
                    {
                        if (transacao.GetObject(idPonto, OpenMode.ForWrite) is CogoPoint cogo)
                            cogo.Layer = request.NomeLayer;
                    }
                    catch (Exception ex)
                    {
                        resultado.Logs.Add(
                            $"Nao foi possivel aplicar layer '{request.NomeLayer}' no ponto {indice}: {ex.Message}");
                    }
                }
            }

            for (int indice = 0; indice < pontosOrdenados.Count; indice++)
            {
                string rotulo = rotulosOrdenados[indice];
                (double, double) chave = ChaveFinaXy(pontosOrdenados[indice]);

                if (!mapaIdPorChave.TryGetValue(chave, out ObjectId idPonto))
                {
                    resultado.Logs.Add($"Sem CogoPoint correspondente para rotulo {rotulo}.");
                    continue;
                }

                try
                {
                    if (transacao.GetObject(idPonto, OpenMode.ForWrite) is CogoPoint cogo)
                    {
                        cogo.RawDescription = rotulo;
                        resultado.TotalRawDescriptionAplicadas++;
                    }
                }
                catch (Exception ex)
                {
                    resultado.Logs.Add($"Erro ao aplicar RawDescription '{rotulo}': {ex.Message}");
                }
            }

            transacao.Commit();
        }
        catch (Exception ex)
        {
            try { transacao.Abort(); } catch { }
            resultado.MensagensErro.Add($"Erro geral ao criar CogoPoints: {ex.Message}");
        }
    }

    /// <summary>
    /// Garante que a layer do destino exista. Criar a layer dentro de sua
    /// propria transacao de escrita evita conflitos com a transacao que
    /// cria os CogoPoints.
    /// </summary>
    private static void GarantirLayerExiste(
        Database db,
        string nomeLayer,
        CriarPontosCruzamentosResultado resultado)
    {
        if (string.IsNullOrWhiteSpace(nomeLayer))
            return;

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            var tabela = (LayerTable)transacao.GetObject(db.LayerTableId, OpenMode.ForRead);

            if (tabela.Has(nomeLayer))
            {
                transacao.Commit();
                return;
            }

            tabela.UpgradeOpen();
            var registro = new LayerTableRecord { Name = nomeLayer };
            tabela.Add(registro);
            transacao.AddNewlyCreatedDBObject(registro, true);
            transacao.Commit();

            resultado.Logs.Add($"Layer '{nomeLayer}' criada.");
        }
        catch (Exception ex)
        {
            try { transacao.Abort(); } catch { }
            resultado.Logs.Add($"Nao foi possivel garantir a layer '{nomeLayer}': {ex.Message}");
        }
    }
}
