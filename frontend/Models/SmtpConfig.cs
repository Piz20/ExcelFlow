namespace ExcelFlow.Models
{
    public class SmtpConfig
    {
        public string? SmtpHost { get; set; }
        public int? SmtpPort { get; set; }
        public string? SmtpFromEmail { get; set; }

        // Méthode pour valider la config (optionnel)
        public bool IsValid()
        {
            return !string.IsNullOrWhiteSpace(SmtpHost)
                && SmtpPort.HasValue && SmtpPort > 0
                && !string.IsNullOrWhiteSpace(SmtpFromEmail);
        }
    }
}
