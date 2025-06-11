using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq; // Added for .Any() and .Union()
using System.Threading.Tasks;

namespace ExcelFlow.Services;

public class SendEmail
{
    private readonly IConfiguration _configuration;
    private readonly string _smtpHost;
    private readonly int _smtpPort;
    public readonly string _fromEmail;

    public SendEmail(IConfiguration configuration)
    {
        _configuration = configuration;

        _smtpHost = _configuration["SmtpSettings:Host"] ?? throw new ArgumentNullException("SmtpSettings:Host is missing in configuration.");
        _smtpPort = int.Parse(_configuration["SmtpSettings:Port"] ?? throw new ArgumentNullException("SmtpSettings:Port is missing in configuration."));
        _fromEmail = _configuration["SmtpSettings:FromEmail"] ?? throw new ArgumentNullException("SmtpSettings:FromEmail is missing in configuration.");
    }

    public string FromEmail => _fromEmail;

    /// <summary>
    /// Sends an email asynchronously with optional attachments to To, Cc, and Bcc recipients.
    /// </summary>
    /// <param name="toRecipients">A list of primary recipient email addresses (visible to all To/Cc).</param>
    /// <param name="ccRecipients">A list of Carbon Copy recipient email addresses (visible to all To/Cc).</param>
    /// <param name="bccRecipients">A list of Blind Carbon Copy recipient email addresses (invisible to other recipients).</param>
    /// <param name="subject">The subject of the email.</param>
    /// <param name="body">The HTML body of the email.</param>
    /// <param name="fromDisplayName">Optional display name for the sender (e.g., "Wafacash Notifications"). Defaults to "Wafacash Mailer".</param>
    /// <param name="attachmentFilePaths">A list of full paths to files to attach to the email. Optional.</param>
    /// <returns>True if the email was sent successfully, false otherwise.</returns>
    public async Task<bool> SendEmailAsync(
        string subject,
        string body,
        List<string>? toRecipients = null,   // New: Optional list for 'To'
        List<string>? ccRecipients = null,   // New: Optional list for 'Cc'
        List<string>? bccRecipients = null,  // New: Optional list for 'Bcc'
        string? fromDisplayName = null,
        List<string>? attachmentFilePaths = null)
    {
        // Ensure at least one recipient list has emails
        if ((toRecipients == null || !toRecipients.Any()) &&
            (ccRecipients == null || !ccRecipients.Any()) &&
            (bccRecipients == null || !bccRecipients.Any()))
        {
            Console.Error.WriteLine("Error: No recipient email addresses provided for To, Cc, or Bcc.");
            return false;
        }

        // Combine all recipients for logging purposes
        var allRecipientsForLogging = (toRecipients ?? Enumerable.Empty<string>())
                                        .Union(ccRecipients ?? Enumerable.Empty<string>())
                                        .Union(bccRecipients ?? Enumerable.Empty<string>())
                                        .ToList();

        try
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(fromDisplayName ?? "Wafacash Mailer", _fromEmail));

            // Add To recipients
            if (toRecipients != null)
            {
                foreach (var email in toRecipients)
                {
                    if (!string.IsNullOrWhiteSpace(email))
                    {
                        message.To.Add(new MailboxAddress("", email.Trim()));
                    }
                }
            }

            // Add Cc recipients
            if (ccRecipients != null)
            {
                foreach (var email in ccRecipients)
                {
                    if (!string.IsNullOrWhiteSpace(email))
                    {
                        message.Cc.Add(new MailboxAddress("", email.Trim()));
                    }
                }
            }

            // Add Bcc recipients
            if (bccRecipients != null)
            {
                foreach (var email in bccRecipients)
                {
                    if (!string.IsNullOrWhiteSpace(email))
                    {
                        message.Bcc.Add(new MailboxAddress("", email.Trim()));
                    }
                }
            }

            // If no valid recipients were added to any field, return false
            if (message.To.Count == 0 && message.Cc.Count == 0 && message.Bcc.Count == 0)
            {
                Console.Error.WriteLine("Error: No valid recipient email addresses found after processing all lists (To, Cc, Bcc).");
                return false;
            }

            message.Subject = subject;

            var bodyBuilder = new BodyBuilder
            {
                HtmlBody = body
            };

            if (attachmentFilePaths != null && attachmentFilePaths.Any()) // Using .Any() for better readability
            {
                foreach (var filePath in attachmentFilePaths)
                {
                    if (File.Exists(filePath))
                    {
                        await bodyBuilder.Attachments.AddAsync(filePath);
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Attachment file not found at path: {filePath}. This attachment will be skipped.");
                    }
                }
            }

            message.Body = bodyBuilder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                await client.ConnectAsync(_smtpHost, _smtpPort, SecureSocketOptions.None);
                await client.AuthenticateAsync(new SaslMechanismGssapi()); // Explicit GSSAPI authentication
                await client.SendAsync(message);
                await client.DisconnectAsync(true);
            }
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error sending email to {string.Join(", ", allRecipientsForLogging)}: {ex.Message}");
            Console.Error.WriteLine($"Full Exception Details: {ex.ToString()}");
            return false;
        }
    }
}