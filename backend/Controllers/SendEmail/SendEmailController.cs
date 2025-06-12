// In your EmailController.cs
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using ExcelFlow.Services;
using ExcelFlow.Models; // Ensure SendEmailRequest is here
using System.Linq; // For .Any() and .Union()

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
                // Assuming fileName already contains the full path from the request.
                // If you intend for "Uploads" to be a subfolder of your application's content root,
                // and fileName is just the file name (e.g., "my_document.pdf"), then
                // string filePath = Path.Combine(_webHostEnvironment.ContentRootPath, "Uploads", fileName);
                string filePath = fileName;

                if (System.IO.File.Exists(filePath))
                {
                    attachmentPaths.Add(filePath);
                }
                else
                {
                    _logger.LogWarning("Attachment file not found: {FileName} at path {FilePath}. Proceeding without this attachment.", fileName, filePath);
                    // You could also return BadRequest here if a missing attachment should fail the request.
                    // return BadRequest($"Attachment file not found: {fileName}");
                }
            }
        }

        // Combine all recipients for logging
        var allRecipientsForLogging = (request.ToRecipients ?? Enumerable.Empty<string>())
                                        .Union(request.CcRecipients ?? Enumerable.Empty<string>())
                                        .Union(request.BccRecipients ?? Enumerable.Empty<string>())
                                        .ToList();

        _logger.LogInformation("Attempting to send email to {Recipients} with subject '{Subject}' from '{FromDisplayName}' with {AttachmentCount} attachments.",
            string.Join(", ", allRecipientsForLogging),
            request.Subject,
            request.FromDisplayName ?? _sendEmail.FromEmail, // Use the FromEmail from the service if FromDisplayName is null
            attachmentPaths?.Count ?? 0);

        // --- CORRECTION START HERE ---
        // Ensure parameters match the SendEmailAsync signature in SendEmail.cs
        bool sent = await _sendEmail.SendEmailAsync(
            toRecipients: request.ToRecipients,   // Pass the List<string> directly
            ccRecipients: request.CcRecipients,   // Pass the List<string> directly
            bccRecipients: request.BccRecipients, // Pass the List<string> directly
            subject: request.Subject,             // Pass the string directly
            body: request.Body,                   // Pass the string directly
            fromDisplayName: request.FromDisplayName, // Pass the string? directly
            attachmentFilePaths: attachmentPaths  // Pass the List<string>? directly
        );
        // --- CORRECTION END HERE ---

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