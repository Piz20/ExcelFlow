using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using ExcelFlow.Services;
using ExcelFlow.Models; // Assure-toi que SendEmailRequest est là
using System.Linq;

[ApiController]
[Route("api/[controller]")]
public class EmailController : ControllerBase
{
    private readonly SendEmail _sendEmail;
    private readonly ILogger<EmailController> _logger;
    private readonly IWebHostEnvironment _webHostEnvironment;

    public EmailController(SendEmail sendEmail, ILogger<EmailController> logger, IWebHostEnvironment webHostEnvironment)
    {
        _sendEmail = sendEmail;
        _logger = logger;
        _webHostEnvironment = webHostEnvironment;
    }

    [HttpPost("send")]
    public async Task<IActionResult> SendEmail([FromBody] SendEmailRequest request)
    {
        if (!ModelState.IsValid)
        {
            _logger.LogWarning("Invalid model state for SendEmail request: {@ModelStateErrors}", ModelState.Values.SelectMany(v => v.Errors));
            return BadRequest(ModelState);
        }

        List<string>? attachmentPaths = null;
        if (request.AttachmentFileNames != null && request.AttachmentFileNames.Any())
        {
            attachmentPaths = new List<string>();
            foreach (var fileName in request.AttachmentFileNames)
            {
                string filePath = fileName; // adapte si besoin (Uploads, etc.)

                if (System.IO.File.Exists(filePath))
                {
                    attachmentPaths.Add(filePath);
                }
                else
                {
                    _logger.LogWarning("Attachment file not found: {FileName} at path {FilePath}. Proceeding without this attachment.", fileName, filePath);
                }
            }
        }

        var allRecipientsForLogging = (request.ToRecipients ?? Enumerable.Empty<string>())
                                        .Union(request.CcRecipients ?? Enumerable.Empty<string>())
                                        .Union(request.BccRecipients ?? Enumerable.Empty<string>())
                                        .ToList();

        _logger.LogInformation("Attempting to send email to {Recipients} with subject '{Subject}' from '{FromDisplayName}' with {AttachmentCount} attachments.",
            string.Join(", ", allRecipientsForLogging),
            request.Subject,
            request.FromDisplayName ?? _sendEmail.FromEmail,
            attachmentPaths?.Count ?? 0);

        // Envoi avec uniquement smtpHost, smtpPort et fromEmail
        bool sent = await _sendEmail.SendEmailAsync(
            subject: request.Subject,
            body: request.Body,
            toRecipients: request.ToRecipients,
            ccRecipients: request.CcRecipients,
            bccRecipients: request.BccRecipients,
            fromDisplayName: request.FromDisplayName,
            attachmentFilePaths: attachmentPaths,
            smtpHost: request.SmtpHost,
            smtpPort: request.SmtpPort,
            smtpFromEmail: request.SmtpFromEmail // ici au lieu de smtpUser/password
        );

        if (sent)
        {
            _logger.LogInformation("Email sent successfully to {Recipients}", string.Join(", ", allRecipientsForLogging));
            return Ok(new { Message = "Email sent successfully!" });
        }
        else
        {
            _logger.LogError("Failed to send email to {Recipients}. Check SendEmail service logs for details.", string.Join(", ", allRecipientsForLogging));
            return StatusCode(500, new { Message = "Failed to send email." });
        }
    }
}
