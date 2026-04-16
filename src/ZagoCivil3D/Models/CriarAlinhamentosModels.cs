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
    /// Requisição para criação de alignments a partir de dois layers:
    /// primeiro as polilinhas horizontais (sentido L-O) ordenadas Norte→Sul,
    /// em seguida as verticais (sentido N-S) ordenadas Oeste→Leste,
    /// com numeração sequencial compartilhada entre os dois grupos.
    /// </summary>
    public sealed class CriarAlinhamentosOrdenadosRequest
    {
        public string Prefixo { get; set; } = "D";
        public string NomeCamadaHorizontais { get; set; } = string.Empty;
        public string NomeCamadaVerticais { get; set; } = string.Empty;
        public string NomeEstiloAlinhamento { get; set; } = string.Empty;
        public string NomeConjuntoRotulosAlinhamento { get; set; } = string.Empty;
        public int NumeroInicial { get; set; } = 1;
        public bool ApagarPolilinhasOriginais { get; set; } = false;

        /// <summary>
        /// Quando verdadeiro (padrão), processa primeiro as polilinhas horizontais
        /// (ordenadas Norte→Sul) e depois as verticais (Oeste→Leste).
        /// Quando falso, inverte: primeiro as verticais e em seguida as horizontais.
        /// </summary>
        public bool HorizontaisPrimeiro { get; set; } = true;
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