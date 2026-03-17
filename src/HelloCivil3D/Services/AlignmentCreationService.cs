// AlignmentCreationService.cs
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using HelloCivil3D.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace HelloCivil3D.Services
{
    public static class AlignmentCreationService
    {
        public static CriarAlinhamentosResultado Executar(
            CivilDocument civilDoc,
            Database db,
            Editor ed,
            CriarAlinhamentosRequest request)
        {
            var resultado = new CriarAlinhamentosResultado();

            using var tr = db.TransactionManager.StartTransaction();

            ObjectId siteId = GetSiteIdByName(civilDoc, tr, request.SiteName);
            if (siteId == ObjectId.Null)
            {
                resultado.MensagensErro.Add($"Site '{request.SiteName}' não encontrado.");
                return resultado;
            }

            ObjectId sourceLayerId = GetLayerIdByName(db, tr, request.SourceLayerName);
            if (sourceLayerId == ObjectId.Null)
            {
                resultado.MensagensErro.Add($"Layer de origem '{request.SourceLayerName}' não encontrada.");
                return resultado;
            }

            ObjectId destinationLayerId = GetLayerIdByName(
                db,
                tr,
                string.IsNullOrWhiteSpace(request.DestinationLayerName)
                    ? request.SourceLayerName
                    : request.DestinationLayerName);

            if (destinationLayerId == ObjectId.Null)
            {
                resultado.MensagensErro.Add($"Layer de destino '{request.DestinationLayerName}' não encontrada.");
                return resultado;
            }

            ObjectId alignmentStyleId = GetAlignmentStyleId(civilDoc, request.AlignmentStyleName);
            if (alignmentStyleId == ObjectId.Null)
            {
                resultado.MensagensErro.Add(
                    $"Estilo de alignment '{request.AlignmentStyleName}' não encontrado.");
                return resultado;
            }

            ObjectId labelSetId = GetAlignmentLabelSetId(civilDoc, request.AlignmentLabelSetName);
            if (labelSetId == ObjectId.Null)
            {
                resultado.MensagensErro.Add(
                    $"Label set de alignment '{request.AlignmentLabelSetName}' não encontrado.");
                return resultado;
            }

            List<ObjectId> polylines = GetPolylinesFromLayer(db, tr, request.SourceLayerName);
            resultado.TotalPolilinhas = polylines.Count;

            if (polylines.Count == 0)
            {
                resultado.MensagensErro.Add(
                    $"Nenhuma polilinha encontrada na layer '{request.SourceLayerName}'.");
                return resultado;
            }

            List<ZonaSiteInfo> zonas = GetFeatureLineZonesForSite(db, tr, siteId);
            if (zonas.Count == 0)
            {
                resultado.MensagensErro.Add(
                    $"Nenhuma Feature Line encontrada associada ao site '{request.SiteName}'.");
                return resultado;
            }

            int numeroAtual = request.NumeroInicial;

            foreach (ObjectId polyId in polylines)
            {
                try
                {
                    if (tr.GetObject(polyId, OpenMode.ForRead) is not Autodesk.AutoCAD.DatabaseServices.Entity entity)
                    {
                        resultado.MensagensErro.Add($"Entidade {polyId.Handle} não pôde ser aberta.");
                        continue;
                    }

                    Extents3d polyExtents;
                    try
                    {
                        polyExtents = entity.GeometricExtents;
                    }
                    catch
                    {
                        resultado.MensagensErro.Add(
                            $"Não foi possível obter extents da polilinha {polyId.Handle}.");
                        continue;
                    }

                    Point3d centro = GetExtentsCenter(polyExtents);

                    ZonaSiteInfo? zona = FindMatchingZone(centro, zonas);
                    if (zona == null)
                    {
                        resultado.TotalSemZona++;
                        resultado.MensagensErro.Add(
                            $"Polilinha {polyId.Handle} não caiu em nenhuma zona do site '{request.SiteName}'.");
                        continue;
                    }

                    string sufixoZona = "." + zona.ZonaId;
                    string nomeBase = $"{request.Prefixo}{numeroAtual:00}{sufixoZona}";
                    string nomeFinal = GetUniqueAlignmentName(civilDoc, tr, nomeBase);

                    var polyOptions = new PolylineOptions
                    {
                        PlineId = polyId,
                        AddCurvesBetweenTangents = false,
                        EraseExistingEntities = request.ApagarPolilinhasOriginais
                    };

                    ObjectId alignmentId = Alignment.Create(
                        civilDoc,
                        polyOptions,
                        nomeFinal,
                        siteId,
                        destinationLayerId,
                        alignmentStyleId,
                        labelSetId);

                    if (alignmentId == ObjectId.Null)
                    {
                        resultado.MensagensErro.Add(
                            $"Falha ao criar alignment para polilinha {polyId.Handle}.");
                        continue;
                    }

                    resultado.TotalCriados++;
                    resultado.NomesCriados.Add(nomeFinal);
                    ed.WriteMessage(
                        $"\n[HelloCivil3D] Alignment criado: {nomeFinal} (zona {zona.ZonaId})");

                    numeroAtual += request.Incremento;
                }
                catch (System.Exception ex)
                {
                    resultado.MensagensErro.Add(
                        $"Erro ao processar polilinha {polyId.Handle}: {ex.Message}");
                }
            }

            tr.Commit();
            return resultado;
        }

        private static List<ObjectId> GetPolylinesFromLayer(
            Database db,
            Transaction tr,
            string layerName)
        {
            var ids = new List<ObjectId>();

            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms =
                (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in ms)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is not Autodesk.AutoCAD.DatabaseServices.Entity ent)
                    continue;

                if (!string.Equals(ent.Layer, layerName, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (ent is Polyline || ent is Polyline2d || ent is Polyline3d)
                    ids.Add(id);
            }

            return ids;
        }

        private static List<ZonaSiteInfo> GetFeatureLineZonesForSite(
            Database db,
            Transaction tr,
            ObjectId siteId)
        {
            var zonas = new List<ZonaSiteInfo>();

            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms =
                (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in ms)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is not FeatureLine fl)
                    continue;

                if (fl.SiteId != siteId)
                    continue;

                string nome = fl.Name ?? string.Empty;
                string zonaId = ExtractZonaId(nome);

                Extents3d extents;
                try
                {
                    extents = fl.GeometricExtents;
                }
                catch
                {
                    continue;
                }

                zonas.Add(new ZonaSiteInfo
                {
                    FeatureLineId = id,
                    FeatureLineName = nome,
                    ZonaId = zonaId,
                    Extents = extents
                });
            }

            return zonas;
        }

        private static string ExtractZonaId(string featureLineName)
        {
            if (string.IsNullOrWhiteSpace(featureLineName))
                return "SEMZONA";

            // Captura o último grupo numérico do nome.
            // Exemplos:
            // ZONA_01 -> 01
            // FL-Z3 -> 3
            // X.Y.12 -> 12
            Match match = Regex.Match(featureLineName, @"(\d+)(?!.*\d)");
            if (match.Success)
                return match.Groups[1].Value;

            return featureLineName.Trim().Replace(" ", "_");
        }

        private static ZonaSiteInfo? FindMatchingZone(
            Point3d centroPolyline,
            List<ZonaSiteInfo> zonas)
        {
            // 1) tenta por containment do ponto central no bbox
            var candidatas = zonas
                .Where(z => ExtentsContainsPoint(z.Extents, centroPolyline))
                .ToList();

            if (candidatas.Count == 1)
                return candidatas[0];

            if (candidatas.Count > 1)
            {
                // fallback: pega a zona cujo centro do extents está mais próximo
                return candidatas
                    .OrderBy(z => DistanceSquared(GetExtentsCenter(z.Extents), centroPolyline))
                    .First();
            }

            // 2) fallback global: zona mais próxima pelo centro do bbox
            return zonas
                .OrderBy(z => DistanceSquared(GetExtentsCenter(z.Extents), centroPolyline))
                .FirstOrDefault();
        }

        private static bool ExtentsContainsPoint(Extents3d ext, Point3d pt)
        {
            return pt.X >= ext.MinPoint.X && pt.X <= ext.MaxPoint.X &&
                   pt.Y >= ext.MinPoint.Y && pt.Y <= ext.MaxPoint.Y &&
                   pt.Z >= ext.MinPoint.Z && pt.Z <= ext.MaxPoint.Z;
        }

        private static Point3d GetExtentsCenter(Extents3d ext)
        {
            return new Point3d(
                (ext.MinPoint.X + ext.MaxPoint.X) * 0.5,
                (ext.MinPoint.Y + ext.MaxPoint.Y) * 0.5,
                (ext.MinPoint.Z + ext.MaxPoint.Z) * 0.5);
        }

        private static double DistanceSquared(Point3d a, Point3d b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            double dz = a.Z - b.Z;
            return dx * dx + dy * dy + dz * dz;
        }

        private static ObjectId GetSiteIdByName(
            CivilDocument civilDoc,
            Transaction tr,
            string siteName)
        {
            foreach (ObjectId siteId in civilDoc.GetSiteIds())
            {
                if (tr.GetObject(siteId, OpenMode.ForRead) is not Autodesk.Civil.DatabaseServices.Site site)
                    continue;

                if (string.Equals(site.Name, siteName, StringComparison.OrdinalIgnoreCase))
                    return siteId;
            }

            return ObjectId.Null;
        }

        private static ObjectId GetLayerIdByName(
            Database db,
            Transaction tr,
            string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName))
                return ObjectId.Null;

            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            return lt.Has(layerName) ? lt[layerName] : ObjectId.Null;
        }

        private static ObjectId GetAlignmentStyleId(
            CivilDocument civilDoc,
            string styleName)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(styleName))
                    return civilDoc.Styles.AlignmentStyles[styleName];
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

        private static ObjectId GetAlignmentLabelSetId(
            CivilDocument civilDoc,
            string labelSetName)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(labelSetName))
                    return civilDoc.Styles.LabelSetStyles.AlignmentLabelSetStyles[labelSetName];
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

        private static string GetUniqueAlignmentName(
            CivilDocument civilDoc,
            Transaction tr,
            string nomeBase)
        {
            if (!AlignmentNameExists(civilDoc, tr, nomeBase))
                return nomeBase;

            int i = 1;
            while (true)
            {
                string candidato = $"{nomeBase}_{i.ToString(CultureInfo.InvariantCulture)}";
                if (!AlignmentNameExists(civilDoc, tr, candidato))
                    return candidato;
                i++;
            }
        }

        private static bool AlignmentNameExists(
            CivilDocument civilDoc,
            Transaction tr,
            string name)
        {
            foreach (ObjectId id in civilDoc.GetAlignmentIds())
            {
                if (tr.GetObject(id, OpenMode.ForRead) is not Alignment alignment)
                    continue;

                if (string.Equals(alignment.Name, name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }
    }
}