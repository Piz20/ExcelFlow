// Fichier : Models/PartnerEmailSenderRequest.cs
using System.ComponentModel.DataAnnotations;

namespace ExcelFlow.Models;

/// <summary>
/// Modèle de requête pour l'envoi d'emails aux partenaires avec des fichiers attachés.
/// </summary>
public class PartnerEmailSenderRequest
{
    [Required(ErrorMessage = "Le chemin du fichier Excel des adresses partenaires est requis.")]
    public string PartnerEmailFilePath { get; set; } = string.Empty; // RENOMMÉ

    [Required(ErrorMessage = "Le chemin du dossier des fichiers générés est requis.")]
    public string GeneratedFilesFolderPath { get; set; } = string.Empty;

    public string Subject { get; set; } = "Rapport Mensuel";
    public string Body { get; set; } = "<p>Cher partenaire,</p><p>Veuillez trouver ci-joint votre rapport mensuel.</p><p>Cordialement,</p><p>Votre Équipe</p>";
    public string? FromDisplayName { get; set; }
}