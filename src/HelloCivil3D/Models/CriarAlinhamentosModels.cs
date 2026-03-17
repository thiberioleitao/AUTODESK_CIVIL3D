// CriarAlinhamentosModels.cs
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace HelloCivil3D.Models
{
    /// <summary>
    /// Dados de entrada preenchidos na UI e consumidos pelo serviço de criação.
    /// </summary>
    public sealed class CriarAlinhamentosRequest
    {
        public string Prefixo { get; set; } = "D";
        public string SourceLayerName { get; set; } = string.Empty;
        public string DestinationLayerName { get; set; } = string.Empty;
        public string SiteName { get; set; } = string.Empty;
        public string AlignmentStyleName { get; set; } = string.Empty;
        public string AlignmentLabelSetName { get; set; } = string.Empty;
        public int NumeroInicial { get; set; } = 1;
        public int Incremento { get; set; } = 1;
        public bool ApagarPolilinhasOriginais { get; set; } = false;
    }

    /// <summary>
    /// Consolida os resultados da execução para exibição no resumo final.
    /// </summary>
    public sealed class CriarAlinhamentosResultado
    {
        public int TotalPolilinhas { get; set; }
        public int TotalCriados { get; set; }
        public int TotalSemZona { get; set; }
        public List<string> NomesCriados { get; } = new();
        public List<string> MensagensErro { get; } = new();
    }

    /// <summary>
    /// Representa uma zona derivada de Feature Line para classificação espacial.
    /// </summary>
    public sealed class ZonaSiteInfo
    {
        public ObjectId FeatureLineId { get; set; } = ObjectId.Null;
        public string FeatureLineName { get; set; } = string.Empty;
        public string ZonaId { get; set; } = string.Empty;
        public Extents3d Extents { get; set; }
    }
}