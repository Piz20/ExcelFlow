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
                var results = await _partnerEmailSender.SendPreparedEmailsAsync(emails, cancellationToken);
                return Ok(results); // Renvoie la liste de résultats au frontend
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Erreur lors de l'envoi des emails.");
                return StatusCode(StatusCodes.Status500InternalServerError, "Erreur interne lors de l'envoi des emails.");
            }
        }




    }
}
