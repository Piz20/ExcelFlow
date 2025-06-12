// Fichier : Services/SendEmail.cs (AUCUN CHANGEMENT PAR RAPPORT À VOTRE ORIGINAL)
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.SignalR; // Ajouté pour IHubContext
using System.Threading; // Ajouté pour CancellationToken
using ExcelFlow.Hubs; // Assurez-vous que le Hub est défini dans votre application

namespace ExcelFlow.Services;

// Assurez-vous que le Hub est défini dans votre application, par exemple:
// public class PartnerFileHub : Hub { }

public class SendEmail
{
    private readonly IConfiguration _configuration;
    private readonly string _smtpHost;
    private readonly int _smtpPort;
    public readonly string _fromEmail;
    private readonly IHubContext<PartnerFileHub>? _hubContext; // Rendre nullable si l'injection est optionnelle

    // Ancien constructeur pour la compatibilité, mais le nouveau sera préféré
    public SendEmail(IConfiguration configuration)
    {
        _configuration = configuration;
        _smtpHost = _configuration["SmtpSettings:Host"] ?? throw new ArgumentNullException("SmtpSettings:Host is missing in configuration.");
        _smtpPort = int.Parse(_configuration["SmtpSettings:Port"] ?? throw new ArgumentNullException("SmtpSettings:Port is missing in configuration."));
        _fromEmail = _configuration["SmtpSettings:FromEmail"] ?? throw new ArgumentNullException("SmtpSettings:FromEmail is missing in configuration.");
        // Le hubContext ne sera pas injecté avec ce constructeur
    }

    // NOUVEAU CONSTRUCTEUR AVEC INJECTION DE HUB
    public SendEmail(IConfiguration configuration, IHubContext<PartnerFileHub> hubContext)
    {
        _configuration = configuration;
        _hubContext = hubContext; // Assigner le hubContext injecté

        _smtpHost = _configuration["SmtpSettings:Host"] ?? throw new ArgumentNullException("SmtpSettings:Host is missing in configuration.");
        _smtpPort = int.Parse(_configuration["SmtpSettings:Port"] ?? throw new ArgumentNullException("SmtpSettings:Port is missing in configuration."));
        _fromEmail = _configuration["SmtpSettings:FromEmail"] ?? throw new ArgumentNullException("SmtpSettings:FromEmail is missing in configuration.");
    }

    public string FromEmail => _fromEmail;

    // Méthode LogAndSend copiée depuis votre exemple
    private async Task LogAndSend(string message, CancellationToken cancellationToken = default)
    {
        string formattedMessage = $"[{DateTime.Now:HH:mm:ss}] {message}";
        Console.WriteLine(formattedMessage); // Toujours logger sur la console du serveur
        if (_hubContext != null)
        {
            await _hubContext.Clients.All.SendAsync("ReceiveMessage", formattedMessage, cancellationToken);
        }
    }
    
    // Surcharge pour les messages d'erreur si nécessaire
    private async Task LogAndSendError(string message, CancellationToken cancellationToken = default)
    {
        string formattedMessage = $"[{DateTime.Now:HH:mm:ss}] ERREUR: {message}";
        Console.Error.WriteLine(formattedMessage); // Toujours logger l'erreur sur la console du serveur
        if (_hubContext != null)
        {
            // Vous pourriez vouloir envoyer un type de message d'erreur différent au client
            await _hubContext.Clients.All.SendAsync("ReceiveErrorMessage", formattedMessage, cancellationToken);
        }
    }


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
    /// <param name="cancellationToken">Token to observe for cancellation requests.</param>
    /// <returns>True if the email was sent successfully, false otherwise.</returns>
    public async Task<bool> SendEmailAsync(
        string subject,
        string body,
        List<string>? toRecipients = null,
        List<string>? ccRecipients = null,
        List<string>? bccRecipients = null,
        string? fromDisplayName = null,
        List<string>? attachmentFilePaths = null,
        CancellationToken cancellationToken = default) // Ajout du CancellationToken
    {
        // Ensure at least one recipient list has emails
        if ((toRecipients == null || !toRecipients.Any()) &&
            (ccRecipients == null || !ccRecipients.Any()) &&
            (bccRecipients == null || !bccRecipients.Any()))
        {
            await LogAndSendError("Erreur: Aucune adresse email de destinataire fournie pour To, Cc ou Bcc.", cancellationToken);
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
                await LogAndSendError("Erreur: Aucune adresse email de destinataire valide trouvée après le traitement de toutes les listes (To, Cc, Bcc).", cancellationToken);
                return false;
            }

            message.Subject = subject;

            var bodyBuilder = new BodyBuilder
            {
                HtmlBody = body // <--- C'est cette ligne qui fait que le corps est traité comme HTML
            };

            if (attachmentFilePaths != null && attachmentFilePaths.Any())
            {
                foreach (var filePath in attachmentFilePaths)
                {
                    cancellationToken.ThrowIfCancellationRequested(); // Vérifier l'annulation
                    if (File.Exists(filePath))
                    {
                        await LogAndSend($"Attaching file: {Path.GetFileName(filePath)}", cancellationToken);
                        await bodyBuilder.Attachments.AddAsync(filePath);
                    }
                    else
                    {
                        await LogAndSend($"Avertissement: Le fichier joint n'a pas été trouvé à l'emplacement: {filePath}. Cette pièce jointe sera ignorée.", cancellationToken);
                    }
                }
            }

            message.Body = bodyBuilder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                await LogAndSend($"Connexion à l'hôte SMTP: {_smtpHost}:{_smtpPort}...", cancellationToken);
                await client.ConnectAsync(_smtpHost, _smtpPort, SecureSocketOptions.None, cancellationToken); // Passer le token d'annulation

                await LogAndSend("Authentification SMTP...", cancellationToken);
                await client.AuthenticateAsync(new SaslMechanismGssapi(), cancellationToken); // Passer le token d'annulation

                await LogAndSend($"Envoi de l'email avec le sujet '{subject}' aux destinataires: {string.Join(", ", allRecipientsForLogging)}...", cancellationToken);
                await client.SendAsync(message, cancellationToken); // Passer le token d'annulation

                await LogAndSend("Déconnexion du serveur SMTP...", cancellationToken);
                await client.DisconnectAsync(true, cancellationToken); // Passer le token d'annulation
            }
            await LogAndSend($"Email envoyé avec succès à {string.Join(", ", allRecipientsForLogging)}.", cancellationToken);
            return true;
        }
        catch (OperationCanceledException)
        {
            await LogAndSendError($"L'opération d'envoi d'email à {string.Join(", ", allRecipientsForLogging)} a été annulée.", cancellationToken);
            return false;
        }
        catch (Exception ex)
        {
            await LogAndSendError($"Erreur lors de l'envoi de l'email à {string.Join(", ", allRecipientsForLogging)}: {ex.Message}", cancellationToken);
            await LogAndSendError($"Détails complets de l'exception: {ex.ToString()}", cancellationToken);
            return false;
        }
    }
}