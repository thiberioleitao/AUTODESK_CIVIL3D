// CriarAlinhamentosModels.cs
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace ZagoCivil3D.Models
{
    /// <summary>
    /// Dados de entrada preenchidos na UI e consumidos pelo serviço de criação.
    /// </summary>
    public sealed class CriarAlinhamentosRequest
    {
        public string Prefixo { get; set; } = "D";
        public string IdentificadorZona { get; set; } = "01";
        public string NomeCamadaOrigem { get; set; } = string.Empty;
        public string NomeCamadaDestino { get; set; } = string.Empty;
        public string NomeEstiloAlinhamento { get; set; } = string.Empty;
        public string NomeConjuntoRotulosAlinhamento { get; set; } = string.Empty;
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

}