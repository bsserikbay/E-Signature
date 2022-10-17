namespace ESignature
{
    public class Signature
    {
        public string? FileName { get; set; }
        public string? SignatureName { get; set; }
        public string? NotBefore { get; set; }
        public string? NotAfter { get; set; }
        public string? AlgorithmName { get; set; }
        public int? Version { get; set; }
        public string? Subject { get; set; }
        public string? Issuer { get; set; }
        public bool Signed { get; set; }

    }
}
