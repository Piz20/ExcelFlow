// Fichier : Controllers/PartnerEmailSenderController.cs
using Microsoft.AspNetCore.Mvc;
using ExcelFlow.Services;
using ExcelFlow.Models; // Assurez-vous que ce namespace est correct pour PartnerEmailSenderRequest
using System.Threading.Tasks;
using System.Threading;
using System.IO; // Pour File.Exists et Directory.Exists

namespace ExcelFlow.Controllers;

[ApiController]
[Route("api/partneremailsender")]
public class PartnerEmailSenderController : ControllerBase
{
    private readonly PartnerEmailSender _partnerEmailSender;

    public PartnerEmailSenderController(PartnerEmailSender partnerEmailSender)
    {
        _partnerEmailSender = partnerEmailSender;
    }

    /// <summary>
    /// Endpoint pour déclencher l'envoi d'emails aux partenaires avec leurs fichiers associés.
    /// </summary>
    /// <param name="request">Les paramètres nécessaires pour l'envoi d'emails (définis dans PartnerEmailSenderRequest).</param>
    /// <param name="cancellationToken">Token d'annulation.</param>
    /// <returns>Un résultat HTTP indiquant le succès ou l'échec de l'opération.</returns>
    [HttpPost("send")]
    [ProducesResponseType(StatusCodes.Status200OK)]
    [ProducesResponseType(StatusCodes.Status400BadRequest)]
    [ProducesResponseType(StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> SendPartnerFiles([FromBody] PartnerEmailSenderRequest request, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid)
        {
            // ModelState.IsValid gérera les erreurs des attributs [Required] dans le DTO
            return BadRequest(ModelState);
        }

        // Vérification basique de l'existence des chemins
        // Utilisez System.IO.Path.Combine pour construire des chemins robustes si les parties proviennent de l'extérieur.
        // Ici, on suppose que les chemins complets sont déjà fournis.
        if (!System.IO.File.Exists(request.PartnerEmailFilePath))
        {
            return BadRequest(new { Message = $"Le fichier Excel des adresses partenaires est introuvable : {request.PartnerEmailFilePath}" });
        }

        if (!System.IO.Directory.Exists(request.GeneratedFilesFolderPath))
        {
            return BadRequest(new { Message = $"Le dossier des fichiers générés est introuvable : {request.GeneratedFilesFolderPath}" });
        }
        
        try
        {
            await _partnerEmailSender.SendEmailsToPartnersWithAttachments(
                partnerEmailFilePath: request.PartnerEmailFilePath,
                generatedFilesFolderPath: request.GeneratedFilesFolderPath,
                subject: request.Subject,
                body: request.Body,
                fromDisplayName: request.FromDisplayName,
                ccRecipients: request.CcRecipients,   // PASSAGE DES ADRESSES CC
                bccRecipients: request.BccRecipients, // PASSAGE DES ADRESSES BCC
                cancellationToken: cancellationToken
            );

            return Ok(new { Message = "Processus d'envoi d'emails aux partenaires initié avec succès. Vérifiez les logs et la console pour le statut." });
        }
        catch (OperationCanceledException)
        {
            return StatusCode(StatusCodes.Status400BadRequest, new { Message = "L'opération d'envoi d'emails a été annulée." });
        }
        catch (FileNotFoundException ex)
        {
            return BadRequest(new { Message = ex.Message });
        }
        catch (InvalidOperationException ex)
        {
            return BadRequest(new { Message = ex.Message });
        }
        catch (Exception ex)
        {
            // Pour des raisons de sécurité, en production, ne pas exposer ex.ToString() directement.
            // Utilisez un logger pour enregistrer les détails complets de l'exception.
            return StatusCode(StatusCodes.Status500InternalServerError, new { Message = $"Une erreur inattendue est survenue : {ex.Message}", Details = ex.ToString() });
        }
    }
}