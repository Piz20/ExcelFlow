namespace ExcelFlow.Models
{
  
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