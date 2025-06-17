// Fichier : Models/PartnerEmailSenderRequest.cs
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;

namespace ExcelFlow.Models;

public class EmailSendRequest
{
    [Required(ErrorMessage = "Le chemin du fichier Excel des adresses partenaires est requis.")]
    public string PartnerEmailFilePath { get; set; } = string.Empty;

    [Required(ErrorMessage = "Le chemin du dossier des fichiers générés est requis.")]
    public string GeneratedFilesFolderPath { get; set; } = string.Empty;

    // Les propriétés SubjectTemplate et BodyTemplate sont supprimées car elles ne sont plus fournies par le client.

    public string? FromDisplayName { get; set; } 

    public List<string>? CcRecipients { get; set; } // Liste d'adresses email pour la copie carbone (Cc)

    public List<string>? BccRecipients { get; set; } // Liste d'adresses email pour la copie carbone invisible (Bcc)

    // Nouvelle propriété : liste des mappings fichiers <-> partenaires
}
