using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.SignalR;
using System.Threading;
using ExcelFlow.Hubs;

namespace ExcelFlow.Services
{
    public class SendEmail
    {
        private readonly IConfiguration _configuration;
        private readonly string _smtpHost;
        private readonly int _smtpPort;
        public readonly string _smtpFromEmail;
        private readonly IHubContext<PartnerFileHub>? _hubContext;

        public SendEmail(IConfiguration configuration)
        {
            _configuration = configuration;
            _smtpHost = _configuration["SmtpSettings:Host"] ?? throw new ArgumentNullException("SmtpSettings:Host is missing in configuration.");
            _smtpPort = int.Parse(_configuration["SmtpSettings:Port"] ?? throw new ArgumentNullException("SmtpSettings:Port is missing in configuration."));
            _smtpFromEmail = _configuration["SmtpSettings:FromEmail"] ?? throw new ArgumentNullException("SmtpSettings:FromEmail is missing in configuration.");
        }

        public SendEmail(IConfiguration configuration, IHubContext<PartnerFileHub> hubContext)
        {
            _configuration = configuration;
            _hubContext = hubContext;

            _smtpHost = _configuration["SmtpSettings:Host"] ?? throw new ArgumentNullException("SmtpSettings:Host is missing in configuration.");
            _smtpPort = int.Parse(_configuration["SmtpSettings:Port"] ?? throw new ArgumentNullException("SmtpSettings:Port is missing in configuration."));
            _smtpFromEmail = _configuration["SmtpSettings:FromEmail"] ?? throw new ArgumentNullException("SmtpSettings:FromEmail is missing in configuration.");
        }

        public string FromEmail => _smtpFromEmail;

        private async Task LogAndSend(string message, CancellationToken cancellationToken = default)
        {
            string formattedMessage = $"[{DateTime.Now:HH:mm:ss}] {message}";
            Console.WriteLine(formattedMessage);
            if (_hubContext != null)
            {
                await _hubContext.Clients.All.SendAsync("ReceiveMessage", formattedMessage, cancellationToken);
            }
        }

        private async Task LogAndSendError(string message, CancellationToken cancellationToken = default)
        {
            string formattedMessage = $"[{DateTime.Now:HH:mm:ss}] ERREUR: {message}";
            Console.Error.WriteLine(formattedMessage);
            if (_hubContext != null)
            {
                await _hubContext.Clients.All.SendAsync("ReceiveErrorMessage", formattedMessage, cancellationToken);
            }
        }

        /// <summary>
        /// Envoie un email de manière asynchrone avec paramètres SMTP optionnels.
        /// Authentification GSSAPI utilisée automatiquement (sans login/mot de passe).
        /// </summary>
        public async Task<bool> SendEmailAsync(
            string subject,
            string body,
            List<string>? toRecipients = null,
            List<string>? ccRecipients = null,
            List<string>? bccRecipients = null,
            string? fromDisplayName = null,
            List<string>? attachmentFilePaths = null,
            string? smtpHost = null,
            int? smtpPort = null,
            string? smtpFromEmail = null,
            CancellationToken cancellationToken = default)
        {
            if ((toRecipients == null || !toRecipients.Any()) &&
                (ccRecipients == null || !ccRecipients.Any()) &&
                (bccRecipients == null || !bccRecipients.Any()))
            {
                await LogAndSendError("Erreur: Aucune adresse email de destinataire fournie pour To, Cc ou Bcc.", cancellationToken);
                return false;
            }

            var allRecipientsForLogging = (toRecipients ?? Enumerable.Empty<string>())
                                             .Union(ccRecipients ?? Enumerable.Empty<string>())
                                             .Union(bccRecipients ?? Enumerable.Empty<string>())
                                             .ToList();

            try
            {
                var message = new MimeMessage();
                string senderEmail = smtpFromEmail ?? _smtpFromEmail;
                message.From.Add(new MailboxAddress(fromDisplayName ?? "Wafacash Mailer", senderEmail));

                if (toRecipients != null)
                    foreach (var email in toRecipients)
                        if (!string.IsNullOrWhiteSpace(email))
                            message.To.Add(new MailboxAddress("", email.Trim()));

                if (ccRecipients != null)
                    foreach (var email in ccRecipients)
                        if (!string.IsNullOrWhiteSpace(email))
                            message.Cc.Add(new MailboxAddress("", email.Trim()));

                if (bccRecipients != null)
                    foreach (var email in bccRecipients)
                        if (!string.IsNullOrWhiteSpace(email))
                            message.Bcc.Add(new MailboxAddress("", email.Trim()));

                if (message.To.Count == 0 && message.Cc.Count == 0 && message.Bcc.Count == 0)
                {
                    await LogAndSendError("Erreur: Aucune adresse email de destinataire valide trouvée.", cancellationToken);
                    return false;
                }

                message.Subject = subject;

                var bodyBuilder = new BodyBuilder { HtmlBody = body };

                if (attachmentFilePaths != null && attachmentFilePaths.Any())
                {
                    foreach (var filePath in attachmentFilePaths)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (File.Exists(filePath))
                        {
                            await LogAndSend($"Attaching file: {Path.GetFileName(filePath)}", cancellationToken);
                            await bodyBuilder.Attachments.AddAsync(filePath);
                        }
                        else
                        {
                            await LogAndSend($"Avertissement: Le fichier joint non trouvé: {filePath}. Pièce jointe ignorée.", cancellationToken);
                        }
                    }
                }

                message.Body = bodyBuilder.ToMessageBody();

                string host = smtpHost ?? _smtpHost;
                int port = smtpPort ?? _smtpPort;
                string FromEmail = smtpFromEmail ?? _smtpFromEmail;

                using (var client = new SmtpClient())
                {
                    client.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;

                    await LogAndSend($"Connexion SMTP: host={host}, port={port}, fromEmail={FromEmail}...", cancellationToken);
                    await client.ConnectAsync(host, port, SecureSocketOptions.StartTls, cancellationToken);

                    await LogAndSend("Authentification via GSSAPI (Kerberos)...", cancellationToken);
                    await client.AuthenticateAsync(new SaslMechanismGssapi(), cancellationToken);

                    await LogAndSend($"Envoi de l'email '{subject}' aux destinataires: {string.Join(", ", allRecipientsForLogging)}...", cancellationToken);
                    await client.SendAsync(message, cancellationToken);

                    await LogAndSend("Déconnexion du serveur SMTP...", cancellationToken);
                    await client.DisconnectAsync(true, cancellationToken);
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
                await LogAndSendError($"Détails complets de l'exception: {ex}", cancellationToken);
                return false;
            }
        }
    }
}
