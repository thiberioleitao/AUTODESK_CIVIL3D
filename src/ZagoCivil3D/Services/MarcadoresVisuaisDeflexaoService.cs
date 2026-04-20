using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Versioning;
using ZagoCivil3D.Models;

namespace ZagoCivil3D.Services;

[SupportedOSPlatform("windows")]
/// <summary>
/// Cria marcadores visuais (circulo ao redor do PI + multileader com seta apontando para
/// o centro) no desenho para identificar rapidamente o resultado do ajuste em cada PI:
///
/// <list type="bullet">
///   <item>Verde: PIs ajustados com sucesso (violacao inicial corrigida).</item>
///   <item>Vermelho: PIs em violacao remanescente ou falha geometrica.</item>
///   <item>Azul: PIs inalterados (ja estavam dentro do limite; apenas circulo, sem label).</item>
/// </list>
///
/// Cada status vai para sua propria layer, permitindo ao usuario controlar a visibilidade
/// via gerenciador de layers. As entidades sao identificadas por XData (RegApp
/// <c>ZAGO_DEFLEXAO</c>) para que a proxima execucao remova apenas os marcadores deste
/// comando, sem interferir em outros elementos na mesma layer.
/// </summary>
public static class MarcadoresVisuaisDeflexaoService
{
    /// <summary>Layer padrao para PIs ajustados com sucesso.</summary>
    public const string LayerPadraoAjustado = "ZAGO_DEFLEXAO_AJUSTADO";

    /// <summary>Layer padrao para PIs em violacao remanescente ou falha.</summary>
    public const string LayerPadraoFalha = "ZAGO_DEFLEXAO_FALHA";

    /// <summary>Layer padrao para PIs inalterados (ja dentro do limite).</summary>
    public const string LayerPadraoInalterado = "ZAGO_DEFLEXAO_INALTERADO";

    /// <summary>Nome do RegApp usado para identificar marcadores criados por este comando.</summary>
    private const string m_xdataAppName = "ZAGO_DEFLEXAO";

    /// <summary>Cor ACI verde (ajustado com sucesso).</summary>
    private const short m_corVerde = 3;

    /// <summary>Cor ACI vermelha (violacao / falha).</summary>
    private const short m_corVermelha = 1;

    /// <summary>Cor ACI azul (inalterado).</summary>
    private const short m_corAzul = 5;

    /// <summary>
    /// Cria marcadores visuais a partir do relatorio de pontos. Retorna a quantidade criada.
    /// </summary>
    public static int Criar(
        Database db,
        Editor editor,
        IReadOnlyList<RelatorioPontoPI> pontos,
        string nomeLayerAjustado,
        string nomeLayerFalha,
        string nomeLayerInalterado,
        double alturaTexto,
        bool limparAnteriores)
    {
        double altura = System.Math.Max(0.1, alturaTexto);
        double raioCirculo = altura * 2.0;

        string layerAjustado = NormalizarNomeLayer(nomeLayerAjustado, LayerPadraoAjustado);
        string layerFalha = NormalizarNomeLayer(nomeLayerFalha, LayerPadraoFalha);
        string layerInalterado = NormalizarNomeLayer(nomeLayerInalterado, LayerPadraoInalterado);

        int criados = 0;

        using Transaction transacao = db.TransactionManager.StartTransaction();

        GarantirRegApp(db, transacao);
        GarantirLayer(db, transacao, layerAjustado, m_corVerde);
        GarantirLayer(db, transacao, layerFalha, m_corVermelha);
        GarantirLayer(db, transacao, layerInalterado, m_corAzul);

        if (limparAnteriores)
            ApagarMarcadoresAnteriores(db, transacao);

        var tabelaBlocos = (BlockTable)transacao.GetObject(db.BlockTableId, OpenMode.ForRead);
        var espacoModelo = (BlockTableRecord)transacao.GetObject(
            tabelaBlocos[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite);

        foreach (RelatorioPontoPI ponto in pontos)
        {
            if (!ponto.PosicaoX.HasValue || !ponto.PosicaoY.HasValue)
                continue;

            EstiloMarcador estilo = ObterEstilo(
                ponto.Status,
                layerAjustado,
                layerFalha,
                layerInalterado);

            var posicaoPI = new Point3d(
                ponto.PosicaoX.Value,
                ponto.PosicaoY.Value,
                ponto.PosicaoZ ?? 0);

            // Circulo ao redor do PI — sempre criado, em todos os status.
            Circle circulo = CriarCirculo(posicaoPI, raioCirculo, estilo);
            espacoModelo.AppendEntity(circulo);
            transacao.AddNewlyCreatedDBObject(circulo, true);
            AnexarXData(circulo, $"CIRC;PI={ponto.IndicePI};FL={ponto.NomeFeatureLine}");
            criados++;

            // Multileader — apenas para estados que o usuario precisa diagnosticar.
            if (estilo.CriarLabel)
            {
                MLeader? mleader = CriarMultileader(
                    db,
                    posicaoPI,
                    altura,
                    estilo,
                    MontarTexto(ponto));

                if (mleader != null)
                {
                    espacoModelo.AppendEntity(mleader);
                    transacao.AddNewlyCreatedDBObject(mleader, true);
                    AnexarXData(mleader, $"LBL;PI={ponto.IndicePI};FL={ponto.NomeFeatureLine}");
                    criados++;
                }
            }
        }

        transacao.Commit();

        editor.WriteMessage(
            $"\n[ZagoCivil3D] {criados} marcador(es) criado(s) " +
            $"(ajustado: '{layerAjustado}', falha: '{layerFalha}', inalterado: '{layerInalterado}').");

        return criados;
    }

    /// <summary>
    /// Mapeia o status do PI para a layer, cor e flag de label correspondentes.
    /// PIs inalterados recebem apenas o circulo (sem label) para reduzir a poluicao visual
    /// quando a feature line tem muitos pontos ja dentro do limite.
    /// </summary>
    private static EstiloMarcador ObterEstilo(
        StatusAjustePI status,
        string layerAjustado,
        string layerFalha,
        string layerInalterado)
    {
        return status switch
        {
            StatusAjustePI.AjustadoComSucesso =>
                new EstiloMarcador(layerAjustado, m_corVerde, CriarLabel: true),

            StatusAjustePI.ViolacaoRemanescente =>
                new EstiloMarcador(layerFalha, m_corVermelha, CriarLabel: true),

            StatusAjustePI.FalhaGeometrica =>
                new EstiloMarcador(layerFalha, m_corVermelha, CriarLabel: true),

            StatusAjustePI.JaDentroDoLimite =>
                new EstiloMarcador(layerInalterado, m_corAzul, CriarLabel: false),

            _ => new EstiloMarcador(layerInalterado, m_corAzul, CriarLabel: false)
        };
    }

    /// <summary>
    /// Monta o conteudo do multileader em duas ou tres linhas.
    /// </summary>
    private static string MontarTexto(RelatorioPontoPI ponto)
    {
        string sufixo = ponto.Status switch
        {
            StatusAjustePI.AjustadoComSucesso => "OK",
            StatusAjustePI.ViolacaoRemanescente => "VIOLA",
            StatusAjustePI.FalhaGeometrica => "FALHA",
            _ => string.Empty
        };

        string linha1 = ponto.NomeFeatureLine;
        string linha2 = $"PI {ponto.IndicePI} [{sufixo}]";

        if (ponto.DeltaZAcumulado > 0)
        {
            string dz = ponto.DeltaZAcumulado.ToString("F3", CultureInfo.InvariantCulture);
            return $"{linha1}\\P{linha2}\\PdZ={dz} m";
        }

        return $"{linha1}\\P{linha2}";
    }

    /// <summary>
    /// Cria um circulo no plano XY ao redor do PI, com a cor e a layer do estilo.
    /// </summary>
    private static Circle CriarCirculo(Point3d centro, double raio, EstiloMarcador estilo)
    {
        return new Circle(centro, Vector3d.ZAxis, raio)
        {
            Layer = estilo.Layer,
            Color = Color.FromColorIndex(ColorMethod.ByAci, estilo.CorACI),
        };
    }

    /// <summary>
    /// Cria um multileader com seta apontando para o centro do PI e texto deslocado no
    /// offset diagonal superior-direito.
    /// </summary>
    private static MLeader CriarMultileader(
        Database db,
        Point3d posicaoPI,
        double alturaTexto,
        EstiloMarcador estilo,
        string conteudo)
    {
        double offset = alturaTexto * 6;
        var posicaoTexto = new Point3d(posicaoPI.X + offset, posicaoPI.Y + offset, posicaoPI.Z);

        var mleader = new MLeader();
        mleader.SetDatabaseDefaults();
        mleader.ContentType = ContentType.MTextContent;

        int idLeader = mleader.AddLeader();
        int idLinha = mleader.AddLeaderLine(idLeader);
        mleader.AddFirstVertex(idLinha, posicaoPI);
        mleader.AddLastVertex(idLinha, posicaoTexto);

        var mtext = new MText();
        mtext.SetDatabaseDefaults(db);
        mtext.Contents = conteudo;
        mtext.TextHeight = alturaTexto;
        mtext.Location = posicaoTexto;
        mtext.Attachment = AttachmentPoint.BottomLeft;
        mleader.MText = mtext;

        mleader.ArrowSize = alturaTexto * 1.2;
        mleader.LandingGap = alturaTexto * 0.3;
        mleader.DoglegLength = alturaTexto * 2.0;

        mleader.Layer = estilo.Layer;
        mleader.Color = Color.FromColorIndex(ColorMethod.ByAci, estilo.CorACI);

        return mleader;
    }

    /// <summary>
    /// Retorna o nome informado quando nao for vazio; caso contrario, o valor padrao.
    /// </summary>
    private static string NormalizarNomeLayer(string informado, string padrao)
    {
        return string.IsNullOrWhiteSpace(informado) ? padrao : informado.Trim();
    }

    /// <summary>
    /// Garante que o RegApp usado pelo XData esteja registrado no desenho.
    /// </summary>
    private static void GarantirRegApp(Database db, Transaction transacao)
    {
        var tabelaRegApp = (RegAppTable)transacao.GetObject(db.RegAppTableId, OpenMode.ForRead);
        if (tabelaRegApp.Has(m_xdataAppName))
            return;

        tabelaRegApp.UpgradeOpen();
        var registro = new RegAppTableRecord { Name = m_xdataAppName };
        tabelaRegApp.Add(registro);
        transacao.AddNewlyCreatedDBObject(registro, true);
    }

    /// <summary>
    /// Cria a layer informada se ela nao existir, com a cor padrao correspondente ao status.
    /// Se a layer ja existe, respeita sua cor atual (nao sobrescreve).
    /// </summary>
    private static void GarantirLayer(Database db, Transaction transacao, string nomeLayer, short corPadrao)
    {
        var tabelaLayers = (LayerTable)transacao.GetObject(db.LayerTableId, OpenMode.ForRead);
        if (tabelaLayers.Has(nomeLayer))
            return;

        tabelaLayers.UpgradeOpen();
        var registro = new LayerTableRecord
        {
            Name = nomeLayer,
            Color = Color.FromColorIndex(ColorMethod.ByAci, corPadrao),
        };
        tabelaLayers.Add(registro);
        transacao.AddNewlyCreatedDBObject(registro, true);
    }

    /// <summary>
    /// Anexa um XData simples a uma entidade para identifica-la como marcador deste comando.
    /// </summary>
    private static void AnexarXData(Autodesk.AutoCAD.DatabaseServices.Entity entidade, string tag)
    {
        var buffer = new ResultBuffer(
            new TypedValue((int)DxfCode.ExtendedDataRegAppName, m_xdataAppName),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, tag));
        entidade.XData = buffer;
    }

    /// <summary>
    /// Apaga todas as entidades do model space que possuem o XData de marcador, independente
    /// da layer atual. Cobre circulos e multileaders criados em execucoes anteriores.
    /// </summary>
    private static void ApagarMarcadoresAnteriores(Database db, Transaction transacao)
    {
        var tabelaBlocos = (BlockTable)transacao.GetObject(db.BlockTableId, OpenMode.ForRead);
        var espacoModelo = (BlockTableRecord)transacao.GetObject(
            tabelaBlocos[BlockTableRecord.ModelSpace],
            OpenMode.ForRead);

        var ids = espacoModelo.Cast<ObjectId>().ToList();

        foreach (ObjectId id in ids)
        {
            if (id.IsNull || id.IsErased)
                continue;

            if (transacao.GetObject(id, OpenMode.ForRead) is not Autodesk.AutoCAD.DatabaseServices.Entity entidade)
                continue;

            ResultBuffer? xdata = entidade.GetXDataForApplication(m_xdataAppName);
            if (xdata == null)
                continue;

            entidade.UpgradeOpen();
            entidade.Erase();
        }
    }

    /// <summary>
    /// Agrupa a layer, a cor ACI e a flag de label a aplicar em um marcador.
    /// </summary>
    private readonly record struct EstiloMarcador(string Layer, short CorACI, bool CriarLabel);
}
