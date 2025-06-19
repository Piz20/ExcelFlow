namespace ExcelFlow.Models
{
    /// <summary>
    /// Représente les données de requête envoyées du frontend au backend pour démarrer l'envoi d'emails.
    /// </summary>
    public class EmailSendRequest
    {
        public string GeneratedFilesFolderPath { get; set; } = string.Empty;
        public string PartnerEmailFilePath { get; set; } = string.Empty;
        public string? FromDisplayName { get; set; } = "WAFACASH CENTRAL AFRICA LIMITED";

        public List<string> CcRecipients { get; set; } = new List<string>();
        public List<string> BccRecipients { get; set; } = new List<string>();

        // Paramètres SMTP optionnels sans authentification
        public string? SmtpHost { get; set; } = null;
        public int? SmtpPort { get; set; } = null;
        public string? SmtpFromEmail { get; set; } = null;
    }

    /// <summary>
    /// Représente les informations d'un partenaire identifié.
    /// </summary>
    public class PartnerInfo
    {
        public string PartnerName { get; set; } = string.Empty; // Nom complet du partenaire
        public List<string> Emails { get; set; } = new List<string>(); // Liste des adresses email du partenaire
        public string SearchableNameFull { get; set; } = string.Empty; // Nom complet optimisé pour la recherche (ex: sans accents, espaces)
        public string? SearchableNameSigle { get; set; } // Sigle ou nom abrégé optimisé pour la recherche (nullable)
    }

    /// <summary>
    /// Représente un résumé d'un email envoyé.
    /// </summary>
    public class SentEmailSummary
    {
        public string FileName { get; set; } = string.Empty; // Nom du fichier (ex: Excel ou PDF) envoyé au partenaire
        public string PartnerName { get; set; } = string.Empty; // Nom du partenaire à qui l'email a été envoyé
        public List<string> RecipientEmails { get; set; } = new List<string>(); // Adresses des destinataires principaux (To)
        public List<string> CcRecipientsSent { get; set; } = new List<string>(); // Adresses en copie carbone (CC)
        public List<string> BccRecipientsSent { get; set; } = new List<string>(); // Adresses en copie carbone invisible (BCC)
    }
}