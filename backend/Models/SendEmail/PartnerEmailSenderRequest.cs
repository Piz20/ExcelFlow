using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
namespace ExcelFlow.Models;
public class PartnerEmailSenderRequest
{
    [Required(ErrorMessage = "Le chemin du fichier Excel des adresses partenaires est requis.")]
    public string PartnerEmailFilePath { get; set; } = string.Empty;

    [Required(ErrorMessage = "Le chemin du dossier des fichiers générés est requis.")]
    public string GeneratedFilesFolderPath { get; set; } = string.Empty;

    public string? FromDisplayName { get; set; }

    public List<string>? CcRecipients { get; set; }

    public List<string>? BccRecipients { get; set; }

    // SMTP optionnel, donc pas de [Required]
    public string? SmtpHost { get; set; }

    [Range(1, 65535, ErrorMessage = "Le port SMTP doit être compris entre 1 et 65535.")]
    public int? SmtpPort { get; set; }

    public string? SmtpFromEmail { get; set; }
}
