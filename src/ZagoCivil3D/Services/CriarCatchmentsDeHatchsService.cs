using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Porta em C# da rotina Dynamo
/// "CRIAR CATCHMENT A PARTIR DAS HATCHS". A rotina le todas as hatches de
/// uma layer, associa cada uma a um MText (ID da subbacia) e a uma polilinha
/// (talvegue), e cria um Catchment no Civil 3D com boundary = contorno da
/// hatch e flow path = talvegue.
///
/// A logica de "ponto dentro do contorno" usa o algoritmo de ray casting em
/// 2D, para nao depender de componentes Dynamo.
/// </summary>
public static class CriarCatchmentsDeHatchsService
{
    /// <summary>
    /// Tolerancia de pequenas perturbacoes XY ao testar "contem ponto" contra
    /// o contorno poligonal da hatch. Evita falsos positivos de bordas.
    /// </summary>
    private const double ToleranciaContornoXY = 1e-6;

    /// <summary>
    /// Executa o fluxo completo: coleta entidades, associa cada hatch ao
    /// MText/talvegue correspondentes e cria os catchments.
    /// </summary>
    public static CriarCatchmentsDeHatchsResultado Executar(
        CivilDocument civilDoc,
        Database db,
        Editor ed,
        CriarCatchmentsDeHatchsRequest request)
    {
        var resultado = new CriarCatchmentsDeHatchsResultado();

        if (string.IsNullOrWhiteSpace(request.NomeLayerHatches))
        {
            resultado.MensagensErro.Add("Layer das hatches nao informada.");
            return resultado;
        }
        if (string.IsNullOrWhiteSpace(request.NomeLayerMTextsIds))
        {
            resultado.MensagensErro.Add("Layer dos MTexts (IDs das subbacias) nao informada.");
            return resultado;
        }
        if (string.IsNullOrWhiteSpace(request.NomeGrupoCatchment))
        {
            resultado.MensagensErro.Add("Nome do Catchment Group nao informado.");
            return resultado;
        }

        // Resolve superficie. O Civil 3D exige ObjectId valido ou
        // ObjectId.Null para Catchment.Create — usamos Null quando o usuario
        // nao selecionou nada.
        ObjectId idSuperficie = ObterIdSuperficiePorNome(civilDoc, db, request.NomeSuperficie);
        if (!idSuperficie.IsNull)
            resultado.Logs.Add($"Superficie de referencia: '{request.NomeSuperficie}'.");
        else if (!string.IsNullOrWhiteSpace(request.NomeSuperficie))
            resultado.Logs.Add($"Superficie '{request.NomeSuperficie}' nao encontrada. Catchments serao criados sem superficie de referencia.");

        // Resolve grupo de catchments (cria se nao existir).
        ObjectId idGrupo = ObterOuCriarGrupoCatchment(db, request.NomeGrupoCatchment, resultado);
        if (idGrupo.IsNull)
        {
            // Mensagem de erro ja foi registrada por ObterOuCriarGrupoCatchment
            return resultado;
        }

        // Resolve estilo.
        ObjectId idEstilo = ObterIdEstiloCatchment(civilDoc, db, request.NomeEstiloCatchment, resultado);
        if (idEstilo.IsNull)
        {
            resultado.MensagensErro.Add(
                "Nao foi encontrado nenhum Catchment Style no desenho. Crie ao menos um estilo antes.");
            return resultado;
        }

        // Colecao das hatches, mtexts e polilinhas a trabalhar. As tres
        // colecoes sao lidas sob uma mesma transacao para garantir snapshot
        // coerente.
        List<HatchBoundary> hatches;
        List<MTextInfo> mtexts;
        List<TalvegueInfo> talvegues;

        using (Transaction transacaoLeitura = db.TransactionManager.StartTransaction())
        {
            hatches = ColetarHatches(transacaoLeitura, db, request.NomeLayerHatches, resultado);
            mtexts = ColetarMTexts(transacaoLeitura, db, request.NomeLayerMTextsIds, resultado);
            talvegues = ColetarTalvegues(transacaoLeitura, db, request.NomeLayerTalvegues, resultado);
            transacaoLeitura.Commit();
        }

        resultado.TotalHatchesEncontradas = hatches.Count;
        resultado.TotalMTextsEncontrados = mtexts.Count;
        resultado.TotalTalveguesEncontrados = talvegues.Count;

        if (hatches.Count == 0)
        {
            resultado.MensagensErro.Add(
                $"Nenhuma hatch encontrada na layer '{request.NomeLayerHatches}'.");
            return resultado;
        }

        ed.WriteMessage(
            $"\n[ZagoCivil3D] Criando catchments a partir de {hatches.Count} hatch(es) " +
            $"(grupo '{request.NomeGrupoCatchment}')...");

        // Cada hatch em sua propria transacao para que um erro isolado nao
        // aborte o lote.
        foreach (HatchBoundary hatch in hatches)
        {
            ProcessarHatch(
                db,
                hatch,
                mtexts,
                talvegues,
                idGrupo,
                idEstilo,
                idSuperficie,
                request,
                resultado);
        }

        return resultado;
    }

    /// <summary>
    /// Processa uma unica hatch: procura o MText de ID, procura o talvegue,
    /// e cria o catchment com flow path.
    /// </summary>
    private static void ProcessarHatch(
        Database db,
        HatchBoundary hatch,
        List<MTextInfo> mtexts,
        List<TalvegueInfo> talvegues,
        ObjectId idGrupo,
        ObjectId idEstilo,
        ObjectId idSuperficie,
        CriarCatchmentsDeHatchsRequest request,
        CriarCatchmentsDeHatchsResultado resultado)
    {
        // Localiza o MText dentro do contorno.
        MTextInfo? mtextDentro = mtexts.FirstOrDefault(m => PontoDentroDoPoligono(hatch.ContornoXY, m.Posicao));
        if (mtextDentro == null)
        {
            resultado.AvisosHatchesIgnoradas.Add(
                $"Hatch {hatch.Handle}: nenhum MText da layer de IDs esta contido no contorno. Hatch ignorada.");
            return;
        }

        string idSubbacia = NormalizarIdSubbacia(mtextDentro.Conteudo);
        if (string.IsNullOrWhiteSpace(idSubbacia))
        {
            resultado.AvisosHatchesIgnoradas.Add(
                $"Hatch {hatch.Handle}: MText encontrado mas com conteudo vazio/invalido. Hatch ignorada.");
            return;
        }

        string nomeCatchment = string.IsNullOrWhiteSpace(request.PrefixoBacia)
            ? idSubbacia
            : request.PrefixoBacia.Trim() + (request.SeparadorNome ?? "-") + idSubbacia;

        // Localiza o talvegue com os dois extremos dentro do contorno (se houver).
        TalvegueInfo? talvegueDentro = null;
        if (request.ConfigurarFlowPath)
        {
            talvegueDentro = talvegues.FirstOrDefault(t =>
                PontoDentroDoPoligono(hatch.ContornoXY, t.Inicio) &&
                PontoDentroDoPoligono(hatch.ContornoXY, t.Fim));
        }

        using Transaction transacao = db.TransactionManager.StartTransaction();

        try
        {
            // Substituir existente, se solicitado.
            if (request.SubstituirExistentes)
            {
                int removidos = RemoverCatchmentExistente(transacao, idGrupo, nomeCatchment);
                if (removidos > 0)
                {
                    resultado.TotalSubstituidos += removidos;
                    resultado.Logs.Add($"Catchment '{nomeCatchment}' existente removido antes de recriar.");
                }
            }

            // Point3dCollection da boundary (fechada). Civil 3D exige
            // primeiro == ultimo nao serem duplicados — entregamos a lista
            // aberta (o Create fecha internamente).
            var boundary = new Point3dCollection();
            foreach (Point3d p in hatch.ContornoXY)
                boundary.Add(new Point3d(p.X, p.Y, 0));

            ObjectId idCatchment = Catchment.Create(
                nomeCatchment,
                idGrupo,
                idEstilo,
                idSuperficie, // pode ser Null
                boundary);

            if (idCatchment.IsNull)
            {
                resultado.MensagensErro.Add(
                    $"Hatch {hatch.Handle}: Catchment.Create retornou ObjectId.Null.");
                transacao.Commit();
                return;
            }

            resultado.TotalCatchmentsCriados++;
            resultado.NomesCriados.Add(nomeCatchment);

            // Define flow path, se houver talvegue.
            // Usa FlowPath.SetPath (API atual) em vez de Catchment.SetFlowPath (obsoleta).
            if (talvegueDentro != null)
            {
                if (transacao.GetObject(idCatchment, OpenMode.ForWrite) is Catchment catchment)
                {
                    var pathPoints = new Point3dCollection();
                    foreach (Point3d p in talvegueDentro.Pontos)
                        pathPoints.Add(p);

                    FlowPath flowPath = catchment.GetFlowPath();
                    flowPath.SetPath(pathPoints);
                    resultado.TotalComFlowPath++;
                    resultado.Logs.Add(
                        $"'{nomeCatchment}': flow path com {talvegueDentro.Pontos.Count} ponto(s).");
                }
            }
            else if (request.ConfigurarFlowPath)
            {
                resultado.Logs.Add(
                    $"'{nomeCatchment}': nenhum talvegue inteiramente contido no contorno. Catchment criado sem flow path.");
            }

            transacao.Commit();
        }
        catch (System.Exception ex)
        {
            resultado.MensagensErro.Add(
                $"Hatch {hatch.Handle} (subbacia '{idSubbacia}'): erro ao criar catchment - {ex.Message}");
        }
    }

    /// <summary>
    /// Remove todos os catchments do grupo com o mesmo nome.
    /// </summary>
    private static int RemoverCatchmentExistente(Transaction transacao, ObjectId idGrupo, string nome)
    {
        if (transacao.GetObject(idGrupo, OpenMode.ForWrite) is not CatchmentGroup grupo)
            return 0;

        ObjectId idExistente = grupo.GetCatchmentId(nome);
        if (idExistente.IsNull || idExistente.IsErased)
            return 0;

        if (transacao.GetObject(idExistente, OpenMode.ForWrite) is Autodesk.AutoCAD.DatabaseServices.DBObject dbObj)
        {
            dbObj.Erase();
            return 1;
        }
        return 0;
    }

    /// <summary>
    /// Coleta todas as hatches da layer informada e extrai o contorno XY
    /// (lista de pontos 2D) de cada uma.
    /// </summary>
    private static List<HatchBoundary> ColetarHatches(
        Transaction transacao,
        Database db,
        string nomeLayer,
        CriarCatchmentsDeHatchsResultado resultado)
    {
        var lista = new List<HatchBoundary>();

        var bloco = (BlockTableRecord)transacao.GetObject(
            SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

        foreach (ObjectId id in bloco)
        {
            if (id.IsNull || id.IsErased) continue;
            if (transacao.GetObject(id, OpenMode.ForRead) is not Hatch hatch) continue;
            if (!string.Equals(hatch.Layer, nomeLayer, StringComparison.OrdinalIgnoreCase)) continue;

            List<Point3d> contorno = ExtrairContornoDaHatch(hatch, resultado);
            if (contorno.Count < 3)
                continue;

            lista.Add(new HatchBoundary
            {
                Handle = hatch.Handle.ToString(),
                ContornoXY = contorno,
            });
        }

        return lista;
    }

    /// <summary>
    /// Extrai um contorno fechado a partir da primeira hatch loop externa.
    /// Em hatches baseadas em polilinha, a loop vem como BulgeVertexCollection;
    /// em hatches baseadas em curvas, a loop vem como Curve2dCollection e nos
    /// aproximamos cada curva por segmentos retos (suficiente para point-in-polygon).
    /// </summary>
    private static List<Point3d> ExtrairContornoDaHatch(Hatch hatch, CriarCatchmentsDeHatchsResultado resultado)
    {
        var pontos = new List<Point3d>();

        try
        {
            int numLoops = hatch.NumberOfLoops;
            for (int i = 0; i < numLoops; i++)
            {
                HatchLoop loop = hatch.GetLoopAt(i);

                // Interessa apenas a loop externa.
                if ((loop.LoopType & HatchLoopTypes.External) == 0
                    && (loop.LoopType & HatchLoopTypes.Outermost) == 0)
                    continue;

                if (loop.IsPolyline)
                {
                    BulgeVertexCollection verts = loop.Polyline;
                    foreach (BulgeVertex bv in verts)
                        pontos.Add(new Point3d(bv.Vertex.X, bv.Vertex.Y, 0));
                }
                else
                {
                    // Aproxima cada curva 2D por seus pontos extremos.
                    foreach (Curve2d curva in loop.Curves)
                    {
                        Point2d p0 = curva.StartPoint;
                        Point2d p1 = curva.EndPoint;
                        if (pontos.Count == 0)
                            pontos.Add(new Point3d(p0.X, p0.Y, 0));
                        pontos.Add(new Point3d(p1.X, p1.Y, 0));
                    }
                }

                if (pontos.Count >= 3)
                    break; // primeira loop externa valida ja serve.
            }
        }
        catch (Exception ex)
        {
            resultado.Logs.Add($"Hatch {hatch.Handle}: falha ao extrair contorno ({ex.Message}).");
        }

        // Remove duplicata final se primeira e ultima forem iguais.
        if (pontos.Count >= 2
            && Math.Abs(pontos[0].X - pontos[^1].X) < ToleranciaContornoXY
            && Math.Abs(pontos[0].Y - pontos[^1].Y) < ToleranciaContornoXY)
        {
            pontos.RemoveAt(pontos.Count - 1);
        }

        return pontos;
    }

    /// <summary>
    /// Coleta todos os MTexts da layer informada, normalizando o conteudo
    /// (remove formatacao Mtext tipica, se houver).
    /// </summary>
    private static List<MTextInfo> ColetarMTexts(
        Transaction transacao,
        Database db,
        string nomeLayer,
        CriarCatchmentsDeHatchsResultado resultado)
    {
        var lista = new List<MTextInfo>();

        var bloco = (BlockTableRecord)transacao.GetObject(
            SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

        foreach (ObjectId id in bloco)
        {
            if (id.IsNull || id.IsErased) continue;

            // Aceita tanto MText quanto DBText (pode haver projetos usando o tipo simples).
            if (transacao.GetObject(id, OpenMode.ForRead) is MText mt
                && string.Equals(mt.Layer, nomeLayer, StringComparison.OrdinalIgnoreCase))
            {
                lista.Add(new MTextInfo
                {
                    Conteudo = LimparFormatacaoMText(mt.Contents ?? mt.Text ?? string.Empty),
                    Posicao = new Point3d(mt.Location.X, mt.Location.Y, 0),
                });
            }
            else if (transacao.GetObject(id, OpenMode.ForRead) is DBText dbText
                     && string.Equals(dbText.Layer, nomeLayer, StringComparison.OrdinalIgnoreCase))
            {
                lista.Add(new MTextInfo
                {
                    Conteudo = dbText.TextString ?? string.Empty,
                    Posicao = new Point3d(dbText.Position.X, dbText.Position.Y, 0),
                });
            }
        }

        return lista;
    }

    /// <summary>
    /// Coleta todas as polilinhas da layer informada e extrai seus pontos
    /// (para usar como flow path).
    /// </summary>
    private static List<TalvegueInfo> ColetarTalvegues(
        Transaction transacao,
        Database db,
        string nomeLayer,
        CriarCatchmentsDeHatchsResultado resultado)
    {
        var lista = new List<TalvegueInfo>();
        if (string.IsNullOrWhiteSpace(nomeLayer))
            return lista;

        var bloco = (BlockTableRecord)transacao.GetObject(
            SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

        foreach (ObjectId id in bloco)
        {
            if (id.IsNull || id.IsErased) continue;

            var entidade = transacao.GetObject(id, OpenMode.ForRead);
            if (entidade is not Autodesk.AutoCAD.DatabaseServices.Entity ent) continue;
            if (!string.Equals(ent.Layer, nomeLayer, StringComparison.OrdinalIgnoreCase)) continue;

            List<Point3d>? pontos = null;

            if (entidade is Polyline poli)
            {
                pontos = new List<Point3d>();
                for (int i = 0; i < poli.NumberOfVertices; i++)
                    pontos.Add(poli.GetPoint3dAt(i));
            }
            else if (entidade is Polyline2d poli2d)
            {
                pontos = new List<Point3d>();
                foreach (ObjectId vid in poli2d)
                {
                    if (transacao.GetObject(vid, OpenMode.ForRead) is Vertex2d v2d)
                        pontos.Add(new Point3d(v2d.Position.X, v2d.Position.Y, v2d.Position.Z));
                }
            }
            else if (entidade is Polyline3d poli3d)
            {
                pontos = new List<Point3d>();
                foreach (ObjectId vid in poli3d)
                {
                    if (transacao.GetObject(vid, OpenMode.ForRead) is PolylineVertex3d v3d)
                        pontos.Add(v3d.Position);
                }
            }
            else if (entidade is Line linha)
            {
                pontos = new List<Point3d> { linha.StartPoint, linha.EndPoint };
            }

            if (pontos == null || pontos.Count < 2)
                continue;

            lista.Add(new TalvegueInfo
            {
                Pontos = pontos,
                Inicio = new Point3d(pontos[0].X, pontos[0].Y, 0),
                Fim = new Point3d(pontos[^1].X, pontos[^1].Y, 0),
            });
        }

        return lista;
    }

    /// <summary>
    /// Localiza o ObjectId de uma superficie pelo nome (case-insensitive).
    /// Retorna ObjectId.Null se nao encontrar.
    /// </summary>
    private static ObjectId ObterIdSuperficiePorNome(
        CivilDocument civilDoc,
        Database db,
        string nome)
    {
        if (string.IsNullOrWhiteSpace(nome))
            return ObjectId.Null;

        using Transaction transacao = db.TransactionManager.StartTransaction();

        foreach (ObjectId idSup in civilDoc.GetSurfaceIds())
        {
            if (transacao.GetObject(idSup, OpenMode.ForRead) is Autodesk.Civil.DatabaseServices.Surface sup
                && string.Equals(sup.Name, nome, StringComparison.OrdinalIgnoreCase))
            {
                transacao.Commit();
                return idSup;
            }
        }

        transacao.Commit();
        return ObjectId.Null;
    }

    /// <summary>
    /// Localiza o grupo de catchment pelo nome, criando-o se nao existir.
    /// </summary>
    private static ObjectId ObterOuCriarGrupoCatchment(
        Database db,
        string nome,
        CriarCatchmentsDeHatchsResultado resultado)
    {
        try
        {
            CatchmentGroupCollection grupos = CatchmentGroupCollection.GetCatchmentGroups(db);

            using Transaction transacao = db.TransactionManager.StartTransaction();

            // Procura grupo existente.
            foreach (ObjectId idGrupo in grupos)
            {
                if (idGrupo.IsNull || idGrupo.IsErased) continue;
                if (transacao.GetObject(idGrupo, OpenMode.ForRead) is CatchmentGroup grupo
                    && string.Equals(grupo.Name, nome, StringComparison.OrdinalIgnoreCase))
                {
                    transacao.Commit();
                    resultado.Logs.Add($"Catchment Group '{nome}' ja existente reutilizado.");
                    return idGrupo;
                }
            }

            transacao.Commit();

            // Nao achou — cria.
            ObjectId idNovo = CatchmentGroup.Create(db, nome);
            if (idNovo.IsNull)
            {
                resultado.MensagensErro.Add($"Falha ao criar Catchment Group '{nome}'.");
                return ObjectId.Null;
            }

            resultado.Logs.Add($"Catchment Group '{nome}' criado.");
            return idNovo;
        }
        catch (Exception ex)
        {
            resultado.MensagensErro.Add($"Erro ao localizar/criar Catchment Group '{nome}': {ex.Message}");
            return ObjectId.Null;
        }
    }

    /// <summary>
    /// Localiza o estilo de catchment pelo nome. Se nome vazio ou nao encontrado,
    /// retorna o primeiro estilo disponivel.
    /// </summary>
    private static ObjectId ObterIdEstiloCatchment(
        CivilDocument civilDoc,
        Database db,
        string nomeEstilo,
        CriarCatchmentsDeHatchsResultado resultado)
    {
        using Transaction transacao = db.TransactionManager.StartTransaction();

        ObjectId primeiro = ObjectId.Null;

        foreach (ObjectId idEst in civilDoc.Styles.CatchmentStyles)
        {
            if (idEst.IsNull || idEst.IsErased) continue;
            if (transacao.GetObject(idEst, OpenMode.ForRead) is not StyleBase estilo) continue;

            if (primeiro.IsNull)
                primeiro = idEst;

            if (!string.IsNullOrWhiteSpace(nomeEstilo)
                && string.Equals(estilo.Name, nomeEstilo, StringComparison.OrdinalIgnoreCase))
            {
                transacao.Commit();
                return idEst;
            }
        }

        transacao.Commit();

        if (!string.IsNullOrWhiteSpace(nomeEstilo) && !primeiro.IsNull)
            resultado.Logs.Add(
                $"Estilo '{nomeEstilo}' nao encontrado. Usando o primeiro estilo disponivel.");

        return primeiro;
    }

    /// <summary>
    /// Testa se um ponto XY esta dentro de um poligono definido por uma lista
    /// de vertices (algoritmo de ray casting). Ignora coordenada Z.
    /// </summary>
    private static bool PontoDentroDoPoligono(List<Point3d> poligono, Point3d ponto)
    {
        if (poligono.Count < 3) return false;

        double x = ponto.X;
        double y = ponto.Y;
        bool dentro = false;

        int j = poligono.Count - 1;
        for (int i = 0; i < poligono.Count; i++)
        {
            double xi = poligono[i].X, yi = poligono[i].Y;
            double xj = poligono[j].X, yj = poligono[j].Y;

            bool intersecta = ((yi > y) != (yj > y))
                              && (x < ((xj - xi) * (y - yi) / (yj - yi + 1e-18)) + xi);
            if (intersecta)
                dentro = !dentro;

            j = i;
        }

        return dentro;
    }

    /// <summary>
    /// Remove formatacao tipica de MText (cores, fontes, quebras) e retorna
    /// apenas o texto visivel. Implementacao conservadora: remove blocos
    /// entre chaves e escapes bem conhecidos.
    /// </summary>
    private static string LimparFormatacaoMText(string bruto)
    {
        if (string.IsNullOrEmpty(bruto)) return string.Empty;

        // Remove codigos \f, \F, \C, \H, \T, \Q, \p, \S, \A etc ate o proximo ; ou espaco.
        var sb = new System.Text.StringBuilder(bruto.Length);
        for (int i = 0; i < bruto.Length; i++)
        {
            char c = bruto[i];
            if (c == '\\' && i + 1 < bruto.Length)
            {
                char pr = bruto[i + 1];
                // Escape de caracter literal (\\, \{, \})
                if (pr == '\\' || pr == '{' || pr == '}')
                {
                    sb.Append(pr);
                    i++;
                    continue;
                }
                // Codigos de formatacao: avanca ate proximo ';' ou espaco.
                int j = i + 1;
                while (j < bruto.Length && bruto[j] != ';' && bruto[j] != ' ')
                    j++;
                if (j < bruto.Length && bruto[j] == ';') j++;
                i = j - 1;
                continue;
            }
            if (c == '{' || c == '}')
                continue;
            sb.Append(c);
        }

        return sb.ToString().Trim();
    }

    /// <summary>
    /// Normaliza o identificador da subbacia (remove espacos extras e
    /// converte para maiuscula).
    /// </summary>
    private static string NormalizarIdSubbacia(string bruto)
    {
        if (string.IsNullOrWhiteSpace(bruto)) return string.Empty;
        return bruto.Trim().ToUpper(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Representa uma hatch ja com o contorno XY extraido.
    /// </summary>
    private sealed class HatchBoundary
    {
        public string Handle { get; set; } = string.Empty;
        public List<Point3d> ContornoXY { get; set; } = new();
    }

    /// <summary>
    /// MText/DBText associado a uma subbacia.
    /// </summary>
    private sealed class MTextInfo
    {
        public string Conteudo { get; set; } = string.Empty;
        public Point3d Posicao { get; set; }
    }

    /// <summary>
    /// Polilinha de talvegue (flow path) com seus pontos.
    /// </summary>
    private sealed class TalvegueInfo
    {
        public List<Point3d> Pontos { get; set; } = new();
        public Point3d Inicio { get; set; }
        public Point3d Fim { get; set; }
    }
}
