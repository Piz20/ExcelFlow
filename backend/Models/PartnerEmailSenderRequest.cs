// Fichier : Models/PartnerEmailSenderRequest.cs
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations; // For [Required]

namespace ExcelFlow.Models;

public class PartnerEmailSenderRequest
{
    [Required(ErrorMessage = "Le chemin du fichier Excel des adresses partenaires est obligatoire.")]
    public string PartnerEmailFilePath { get; set; } = string.Empty;

    [Required(ErrorMessage = "Le chemin du dossier des fichiers générés est obligatoire.")]
    public string GeneratedFilesFolderPath { get; set; } = string.Empty;

    [Required(ErrorMessage = "Le sujet de l'email est obligatoire.")]
    public string Subject { get; set; } = string.Empty;

    [Required(ErrorMessage = "Le corps de l'email est obligatoire.")]
    public string Body { get; set; } = string.Empty;

    public string? FromDisplayName { get; set; }

    // AJOUTÉ : Listes pour Cc et Bcc
    public List<string>? CcRecipients { get; set; } = new List<string>(); // Initialize to avoid null reference if not provided
    public List<string>? BccRecipients { get; set; } = new List<string>(); // Initialize to avoid null reference if not provided
}