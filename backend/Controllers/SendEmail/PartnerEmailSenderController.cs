using Microsoft.AspNetCore.Mvc;
using ExcelFlow.Services;
using ExcelFlow.Models;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using Microsoft.AspNetCore.Http;
using System;
using System.IO;
using System.Collections.Generic;

namespace ExcelFlow.Controllers
{
    [ApiController]
    [Route("api/partneremailsender")]
    public class PartnerEmailSenderController : ControllerBase
    {
        private readonly IPartnerEmailSender _partnerEmailSender;
        private readonly ILogger<PartnerEmailSenderController> _logger;

        public PartnerEmailSenderController(IPartnerEmailSender partnerEmailSender, ILogger<PartnerEmailSenderController> logger)
        {
            _partnerEmailSender = partnerEmailSender;
            _logger = logger;
        }

        /// <summary>
        /// Prépare les emails à envoyer (lecture du fichier Excel, génération des emails avec pièces jointes, etc.)
        /// </summary>
        [HttpPost("prepare")]
        public async Task<IActionResult> PrepareEmails([FromBody] PrepareEmailRequest request, CancellationToken cancellationToken)
        {
            try
            {
                var emails = await _partnerEmailSender.PrepareCompleteEmailsAsync(
                    partnerEmailFilePath: request.PartnerExcelPath,
                    generatedFilesFolderPath: request.GeneratedFilesFolder,
                    smtpFromEmail: request.SmtpFromEmail,
                    smtpHost: request.SmtpHost,
                    smtpPort: request.SmtpPort ?? 587,
                    fromDisplayName: request.FromDisplayName,
                    ccRecipients: request.CcRecipients,
                    bccRecipients: request.BccRecipients,
                    cancellationToken: cancellationToken
                );

                return Ok(emails);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Erreur lors de la préparation des emails.");
                return StatusCode(StatusCodes.Status500InternalServerError, "Erreur interne lors de la préparation des emails.");
            }
        }

        /// <summary>
        /// Envoie les emails déjà préparés.
        /// </summary>
        [HttpPost("send")]
        public async Task<IActionResult> SendEmails([FromBody] List<EmailToSend> emails, CancellationToken cancellationToken)
        {
            try
            {
                await _partnerEmailSender.SendPreparedEmailsAsync(emails, cancellationToken);
                return Ok("Emails envoyés avec succès.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Erreur lors de l'envoi des emails.");
                return StatusCode(StatusCodes.Status500InternalServerError, "Erreur interne lors de l'envoi des emails.");
            }
        }

        /// <summary>
        /// Envoie les fichiers partenaires directement en une seule opération.
        /// </summary>
        [HttpPost("send-partner-files")]
        public async Task<IActionResult> SendPartnerFiles([FromBody] PartnerEmailSenderRequest request, CancellationToken cancellationToken)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (!System.IO.File.Exists(request.PartnerEmailFilePath))
            {
                return BadRequest(new { Message = $"Le fichier Excel des adresses partenaires est introuvable : {request.PartnerEmailFilePath}" });
            }

            if (!Directory.Exists(request.GeneratedFilesFolderPath))
            {
                return BadRequest(new { Message = $"Le dossier des fichiers générés est introuvable : {request.GeneratedFilesFolderPath}" });
            }

            if (string.IsNullOrWhiteSpace(request.SmtpHost))
            {
                return BadRequest(new { Message = "Le champ smtpHost est requis et ne peut pas être nul ou vide." });
            }

            if (string.IsNullOrWhiteSpace(request.SmtpFromEmail))
            {
                return BadRequest(new { Message = "Le champ smtpFromEmail est requis et ne peut pas être nul ou vide." });
            }

            if (string.IsNullOrWhiteSpace(request.FromDisplayName))
            {
                return BadRequest(new { Message = "Le champ fromDisplayName est requis et ne peut pas être nul ou vide." });
            }

            try
            {
                await _partnerEmailSender.SendEmailsToPartnersWithAttachments(
                    partnerEmailFilePath: request.PartnerEmailFilePath,
                    generatedFilesFolderPath: request.GeneratedFilesFolderPath,
                    smtpFromEmail: request.SmtpFromEmail!,
                    smtpHost: request.SmtpHost!,
                    smtpPort: request.SmtpPort ?? 587,
                    fromDisplayName: request.FromDisplayName!,
                    ccRecipients: request.CcRecipients ?? new List<string>(),
                    bccRecipients: request.BccRecipients ?? new List<string>(),
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
                return StatusCode(StatusCodes.Status500InternalServerError, new { Message = $"Une erreur inattendue est survenue : {ex.Message}", Details = ex.ToString() });
            }
        }
    }
}
