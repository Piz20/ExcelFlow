namespace ExcelFlow.Models
{
    public class AppConfig
    {
        public SmtpConfig Smtp { get; set; } = new();
        public GenerationDefaults Generation { get; set; } = new();

        public SendEmailDefaults SendEmail { get; set; } = new();
    }

    public class SmtpConfig
    {
        public string? SmtpHost { get; set; }
        public int? SmtpPort { get; set; }
        public string? SmtpFromEmail { get; set; }

        // Méthode de validation de la config SMTP
        public bool IsValid()
        {
            return !string.IsNullOrWhiteSpace(SmtpHost)
                && SmtpPort.HasValue && SmtpPort > 0
                && !string.IsNullOrWhiteSpace(SmtpFromEmail);
        }
    }

    public class GenerationDefaults
    {
        public string? SourcePath { get; set; }
        public string? TemplatePath { get; set; }
        public string? OutputDir { get; set; }
    }

    public class SendEmailDefaults
    {
        public string? PartnerEmailFilePath { get; set; }
        public string? GeneratedFilesFolderPath { get; set; }
        public string? FromDisplayName { get; set; }
        public string? CcRecipients { get; set; }
        public string? BccRecipients { get; set; }
    }
}
