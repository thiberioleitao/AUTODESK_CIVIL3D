// AlignmentCreationService.cs
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using ZagoCivil3D.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ZagoCivil3D.Services
{
    /// <summary>
    /// Concentra as regras de criação de alignments a partir de polilinhas,
    /// incluindo validações, associação por zona e convenção de nomes.
    /// </summary>
    public static class AlignmentCreationService
    {
        /// <summary>
        /// Executa o processamento completo dentro de uma única transação,
        /// garantindo consistência entre leitura e criação das entidades.
        /// </summary>
        public static CriarAlinhamentosResultado Executar(
            CivilDocument civilDoc,
            Database db,
            Editor ed,
            CriarAlinhamentosRequest request)
        {
            var resultado = new CriarAlinhamentosResultado();

            using var transacao = db.TransactionManager.StartTransaction();

            ObjectId idCamadaOrigem = ObterIdCamadaPorNome(db, transacao, request.NomeCamadaOrigem);
            if (idCamadaOrigem == ObjectId.Null)
            {
                resultado.MensagensErro.Add($"Layer de origem '{request.NomeCamadaOrigem}' não encontrada.");
                return resultado;
            }

            ObjectId idEstiloAlinhamento = ObterIdEstiloAlinhamento(civilDoc, request.NomeEstiloAlinhamento);
            if (idEstiloAlinhamento == ObjectId.Null)
            {
                resultado.MensagensErro.Add(
                    $"Estilo de alignment '{request.NomeEstiloAlinhamento}' não encontrado.");
                return resultado;
            }

            ObjectId idConjuntoRotulos = ObterIdConjuntoRotulosAlinhamento(civilDoc, request.NomeConjuntoRotulosAlinhamento);
            if (idConjuntoRotulos == ObjectId.Null)
            {
                resultado.MensagensErro.Add(
                    $"Label set de alignment '{request.NomeConjuntoRotulosAlinhamento}' não encontrado.");
                return resultado;
            }

            List<ObjectId> polilinhas = ObterPolilinhasDaCamada(db, transacao, request.NomeCamadaOrigem);
            resultado.TotalPolilinhas = polilinhas.Count;

            if (polilinhas.Count == 0)
            {
                resultado.MensagensErro.Add(
                    $"Nenhuma polilinha encontrada na layer '{request.NomeCamadaOrigem}'.");
                return resultado;
            }


            int numeroAtual = request.NumeroInicial;

            foreach (ObjectId idPolilinha in polilinhas)
            {
                try
                {
                    if (transacao.GetObject(idPolilinha, OpenMode.ForRead) is not Autodesk.AutoCAD.DatabaseServices.Entity)
                    {
                        resultado.MensagensErro.Add($"Entidade {idPolilinha.Handle} não pôde ser aberta.");
                        continue;
                    }

                    string nomeBase = $"{request.Prefixo}{numeroAtual:00}";
                    string nomeFinal = ObterNomeAlinhamentoUnico(civilDoc, transacao, nomeBase);

                    var opcoesPolilinha = new PolylineOptions
                    {
                        PlineId = idPolilinha,
                        AddCurvesBetweenTangents = false,
                        EraseExistingEntities = request.ApagarPolilinhasOriginais
                    };

                    ObjectId idAlinhamento = Alignment.Create(
                        civilDoc,
                        opcoesPolilinha,
                        nomeFinal,
                        ObjectId.Null,
                        idCamadaOrigem,
                        idEstiloAlinhamento,
                        idConjuntoRotulos);

                    if (idAlinhamento == ObjectId.Null)
                    {
                        resultado.MensagensErro.Add(
                            $"Falha ao criar alignment para polilinha {idPolilinha.Handle}.");
                        continue;
                    }

                    resultado.TotalCriados++;
                    resultado.NomesCriados.Add(nomeFinal);
                    ed.WriteMessage($"\n[ZagoCivil3D] Alignment criado: {nomeFinal}");

                    numeroAtual += 1;
                }
                catch (System.Exception ex)
                {
                    resultado.MensagensErro.Add(
                        $"Erro ao processar polilinha {idPolilinha.Handle}: {ex.Message}");
                }
            }

            transacao.Commit();
            return resultado;
        }

        /// <summary>
        /// Executa criação de alignments a partir de dois layers: primeiro processa
        /// as polilinhas horizontais (sentido L-O) ordenadas Norte→Sul (maior Y primeiro),
        /// em seguida as verticais (sentido N-S) ordenadas Oeste→Leste (menor X primeiro).
        /// A numeração é sequencial e compartilhada entre os dois grupos.
        /// </summary>
        public static CriarAlinhamentosResultado ExecutarOrdenado(
            CivilDocument civilDoc,
            Database db,
            Editor ed,
            CriarAlinhamentosOrdenadosRequest request)
        {
            var resultado = new CriarAlinhamentosResultado();

            using var transacao = db.TransactionManager.StartTransaction();

            ObjectId idCamadaHorizontais = ObterIdCamadaPorNome(db, transacao, request.NomeCamadaHorizontais);
            if (idCamadaHorizontais == ObjectId.Null)
            {
                resultado.MensagensErro.Add(
                    $"Layer de horizontais '{request.NomeCamadaHorizontais}' não encontrada.");
                return resultado;
            }

            ObjectId idCamadaVerticais = ObterIdCamadaPorNome(db, transacao, request.NomeCamadaVerticais);
            if (idCamadaVerticais == ObjectId.Null)
            {
                resultado.MensagensErro.Add(
                    $"Layer de verticais '{request.NomeCamadaVerticais}' não encontrada.");
                return resultado;
            }

            ObjectId idEstilo = ObterIdEstiloAlinhamento(civilDoc, request.NomeEstiloAlinhamento);
            if (idEstilo == ObjectId.Null)
            {
                resultado.MensagensErro.Add(
                    $"Estilo de alignment '{request.NomeEstiloAlinhamento}' não encontrado.");
                return resultado;
            }

            ObjectId idConjuntoRotulos =
                ObterIdConjuntoRotulosAlinhamento(civilDoc, request.NomeConjuntoRotulosAlinhamento);
            if (idConjuntoRotulos == ObjectId.Null)
            {
                resultado.MensagensErro.Add(
                    $"Label set de alignment '{request.NomeConjuntoRotulosAlinhamento}' não encontrado.");
                return resultado;
            }

            List<ObjectId> horizontais = ObterPolilinhasDaCamada(db, transacao, request.NomeCamadaHorizontais);
            List<ObjectId> verticais = ObterPolilinhasDaCamada(db, transacao, request.NomeCamadaVerticais);

            resultado.TotalPolilinhas = horizontais.Count + verticais.Count;

            if (resultado.TotalPolilinhas == 0)
            {
                resultado.MensagensErro.Add(
                    $"Nenhuma polilinha encontrada nas layers '{request.NomeCamadaHorizontais}' e '{request.NomeCamadaVerticais}'.");
                return resultado;
            }

            List<ObjectId> horizontaisOrdenadas = OrdenarPorChave(
                horizontais,
                transacao,
                ext => ext.MaxPoint.Y,
                descendente: true);

            List<ObjectId> verticaisOrdenadas = OrdenarPorChave(
                verticais,
                transacao,
                ext => ext.MinPoint.X,
                descendente: false);

            int numeroAtual = request.NumeroInicial;

            // Monta a ordem de processamento conforme a escolha do usuário.
            // Cada tupla: (polilinhas ordenadas, id da layer destino, rótulo informativo).
            var etapas = new List<(List<ObjectId> Lista, ObjectId IdCamada, string Rotulo)>();
            if (request.HorizontaisPrimeiro)
            {
                etapas.Add((horizontaisOrdenadas, idCamadaHorizontais, "horizontal (N→S)"));
                etapas.Add((verticaisOrdenadas, idCamadaVerticais, "vertical (O→L)"));
            }
            else
            {
                etapas.Add((verticaisOrdenadas, idCamadaVerticais, "vertical (O→L)"));
                etapas.Add((horizontaisOrdenadas, idCamadaHorizontais, "horizontal (N→S)"));
            }

            foreach (var etapa in etapas)
            {
                numeroAtual = CriarAlignmentsDaLista(
                    civilDoc,
                    transacao,
                    ed,
                    etapa.Lista,
                    request.Prefixo,
                    etapa.IdCamada,
                    idEstilo,
                    idConjuntoRotulos,
                    request.ApagarPolilinhasOriginais,
                    numeroAtual,
                    etapa.Rotulo,
                    resultado);
            }

            transacao.Commit();
            return resultado;
        }

        /// <summary>
        /// Ordena polilinhas por uma chave geométrica derivada dos extents.
        /// Polilinhas sem extents válidos ficam no final da lista para não quebrar o fluxo.
        /// </summary>
        private static List<ObjectId> OrdenarPorChave(
            List<ObjectId> polilinhas,
            Transaction transacao,
            Func<Extents3d, double> seletorChave,
            bool descendente)
        {
            var comChave = new List<(ObjectId Id, double Chave, bool Valido)>();

            foreach (ObjectId id in polilinhas)
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is not Autodesk.AutoCAD.DatabaseServices.Entity entidade)
                {
                    comChave.Add((id, 0.0, false));
                    continue;
                }

                try
                {
                    Extents3d extents = entidade.GeometricExtents;
                    comChave.Add((id, seletorChave(extents), true));
                }
                catch
                {
                    comChave.Add((id, 0.0, false));
                }
            }

            IEnumerable<(ObjectId Id, double Chave, bool Valido)> ordenados = descendente
                ? comChave.Where(x => x.Valido).OrderByDescending(x => x.Chave)
                : comChave.Where(x => x.Valido).OrderBy(x => x.Chave);

            var resultado = ordenados.Select(x => x.Id).ToList();
            resultado.AddRange(comChave.Where(x => !x.Valido).Select(x => x.Id));
            return resultado;
        }

        /// <summary>
        /// Cria alignments para cada polilinha da lista, usando o numerador sequencial
        /// compartilhado. Retorna o próximo número disponível após o processamento.
        /// </summary>
        private static int CriarAlignmentsDaLista(
            CivilDocument civilDoc,
            Transaction transacao,
            Editor ed,
            List<ObjectId> polilinhas,
            string prefixo,
            ObjectId idCamadaDestino,
            ObjectId idEstilo,
            ObjectId idConjuntoRotulos,
            bool apagarOriginais,
            int numeroInicial,
            string rotulo,
            CriarAlinhamentosResultado resultado)
        {
            int numeroAtual = numeroInicial;

            foreach (ObjectId idPolilinha in polilinhas)
            {
                try
                {
                    if (transacao.GetObject(idPolilinha, OpenMode.ForRead)
                        is not Autodesk.AutoCAD.DatabaseServices.Entity)
                    {
                        resultado.MensagensErro.Add(
                            $"Entidade {idPolilinha.Handle} não pôde ser aberta ({rotulo}).");
                        continue;
                    }

                    string nomeBase = $"{prefixo}{numeroAtual:00}";
                    string nomeFinal = ObterNomeAlinhamentoUnico(civilDoc, transacao, nomeBase);

                    var opcoesPolilinha = new PolylineOptions
                    {
                        PlineId = idPolilinha,
                        AddCurvesBetweenTangents = false,
                        EraseExistingEntities = apagarOriginais
                    };

                    ObjectId idAlinhamento = Alignment.Create(
                        civilDoc,
                        opcoesPolilinha,
                        nomeFinal,
                        ObjectId.Null,
                        idCamadaDestino,
                        idEstilo,
                        idConjuntoRotulos);

                    if (idAlinhamento == ObjectId.Null)
                    {
                        resultado.MensagensErro.Add(
                            $"Falha ao criar alignment para polilinha {idPolilinha.Handle} ({rotulo}).");
                        continue;
                    }

                    resultado.TotalCriados++;
                    resultado.NomesCriados.Add(nomeFinal);
                    ed.WriteMessage($"\n[ZagoCivil3D] Alignment criado: {nomeFinal} ({rotulo})");

                    numeroAtual += 1;
                }
                catch (System.Exception ex)
                {
                    resultado.MensagensErro.Add(
                        $"Erro ao processar polilinha {idPolilinha.Handle} ({rotulo}): {ex.Message}");
                }
            }

            return numeroAtual;
        }

        private static List<ObjectId> ObterPolilinhasDaCamada(
            Database db,
            Transaction transacao,
            string nomeCamada)
        {
            var ids = new List<ObjectId>();

            var tabelaBlocos = (BlockTable)transacao.GetObject(db.BlockTableId, OpenMode.ForRead);
            var espacoModelo =
                (BlockTableRecord)transacao.GetObject(tabelaBlocos[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in espacoModelo)
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is not Autodesk.AutoCAD.DatabaseServices.Entity entidade)
                    continue;

                if (!string.Equals(entidade.Layer, nomeCamada, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (entidade is Polyline || entidade is Polyline2d || entidade is Polyline3d)
                    ids.Add(id);
            }

            return ids;
        }

        private static ObjectId ObterIdCamadaPorNome(
            Database db,
            Transaction transacao,
            string nomeCamada)
        {
            if (string.IsNullOrWhiteSpace(nomeCamada))
                return ObjectId.Null;

            var tabelaCamadas = (LayerTable)transacao.GetObject(db.LayerTableId, OpenMode.ForRead);
            return tabelaCamadas.Has(nomeCamada) ? tabelaCamadas[nomeCamada] : ObjectId.Null;
        }

        private static ObjectId ObterIdEstiloAlinhamento(
            CivilDocument civilDoc,
            string nomeEstilo)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(nomeEstilo))
                    return civilDoc.Styles.AlignmentStyles[nomeEstilo];
            }
            catch
            {
                // segue para fallback
            }

            try
            {
                if (civilDoc.Styles.AlignmentStyles.Count > 0)
                    return civilDoc.Styles.AlignmentStyles[0];
            }
            catch
            {
            }

            return ObjectId.Null;
        }

        private static ObjectId ObterIdConjuntoRotulosAlinhamento(
            CivilDocument civilDoc,
            string nomeConjuntoRotulos)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(nomeConjuntoRotulos))
                    return civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles[nomeConjuntoRotulos];
            }
            catch
            {
                // segue para fallback
            }

            try
            {
                if (civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles.Count > 0)
                    return civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles[0];
            }
            catch
            {
            }

            return ObjectId.Null;
        }

        private static string ObterNomeAlinhamentoUnico(
            CivilDocument civilDoc,
            Transaction transacao,
            string nomeBase)
        {
            if (!NomeAlinhamentoExiste(civilDoc, transacao, nomeBase))
                return nomeBase;

            int i = 1;
            while (true)
            {
                string candidato = $"{nomeBase}_{i.ToString(CultureInfo.InvariantCulture)}";
                if (!NomeAlinhamentoExiste(civilDoc, transacao, candidato))
                    return candidato;
                i++;
            }
        }

        private static bool NomeAlinhamentoExiste(
            CivilDocument civilDoc,
            Transaction transacao,
            string name)
        {
            foreach (ObjectId id in civilDoc.GetAlignmentIds())
            {
                if (transacao.GetObject(id, OpenMode.ForRead) is not Alignment alignment)
                    continue;

                if (string.Equals(alignment.Name, name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }
    }
}