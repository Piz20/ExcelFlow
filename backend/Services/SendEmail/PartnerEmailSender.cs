// Fichier : Services/PartnerEmailSender.cs
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.SignalR;
using System.Threading;
using ExcelFlow.Hubs;
using ExcelFlow.Utilities;
using ExcelFlow.Models; // Pour PartnerInfo, SentEmailSummary, EmailData, et ProgressUpdate

namespace ExcelFlow.Services;

public class PartnerEmailSender : IPartnerEmailSender
{
    private readonly SendEmail _sendEmailService;
    private readonly IHubContext<PartnerFileHub> _hubContext;
    private readonly EmailDataExtractor _emailDataExtractor;
    private readonly PartnerExcelReader _partnerExcelReader;
    private readonly EmailContentBuilder _emailContentBuilder;

    public PartnerEmailSender(
        SendEmail sendEmailService,
        IHubContext<PartnerFileHub> hubContext,
        EmailDataExtractor emailDataExtractor,
        PartnerExcelReader partnerExcelReader,
        EmailContentBuilder emailContentBuilder)
    {
        _sendEmailService = sendEmailService;
        _hubContext = hubContext;
        _emailDataExtractor = emailDataExtractor;
        _partnerExcelReader = partnerExcelReader;
        _emailContentBuilder = emailContentBuilder;
    }

    // Helper pour envoyer des messages de log généraux (reste inchangé, utilise ReceiveMessage)
    private async Task LogAndSend(string message, CancellationToken cancellationToken = default)
    {
        string formattedMessage = $"[{DateTime.Now:HH:mm:ss}] {message}";
        Console.WriteLine(formattedMessage);
        if (_hubContext != null)
        {
            await _hubContext.Clients.All.SendAsync("ReceiveMessage", formattedMessage, cancellationToken);
        }
    }

    // Helper pour envoyer des messages d'erreur (reste inchangé, utilise ReceiveErrorMessage)
    private async Task LogAndSendError(string message, CancellationToken cancellationToken = default)
    {
        string formattedMessage = $"[{DateTime.Now:HH:mm:ss}] ERREUR: {message}";
        Console.Error.WriteLine(formattedMessage);
        if (_hubContext != null)
        {
            await _hubContext.Clients.All.SendAsync("ReceiveErrorMessage", formattedMessage, cancellationToken);
        }
    }





    public async Task<List<EmailToSend>> PrepareCompleteEmailsAsync(
    string partnerEmailFilePath,
    string generatedFilesFolderPath,
    string? smtpFromEmail,
    string? smtpHost,
    int? smtpPort,
    string? fromDisplayName,
    List<string>? ccRecipients,
    List<string>? bccRecipients,
    CancellationToken cancellationToken)
    {
        var partners = await Task.Run(() =>
            _partnerExcelReader.ReadPartnersFromExcel(partnerEmailFilePath, cancellationToken), cancellationToken);

        var allGeneratedFiles = Directory.GetFiles(generatedFilesFolderPath).ToList();
        var preparedEmails = new List<EmailToSend>();
        var ignoredFilesDetails = new List<(string FileName, string Reason)>();

        for (int i = 0; i < allGeneratedFiles.Count; i++)
        {
            var filePath = allGeneratedFiles[i];

            await LogAndSend(
                $"\n\n\n\n📄 Traitement du fichier {i + 1}/{allGeneratedFiles.Count} : {Path.GetFileName(filePath)}",
                cancellationToken);

            var emailData = await _emailDataExtractor.ExtractEmailDataFromAttachment(filePath, cancellationToken);
            if (emailData == null)
            {
                ignoredFilesDetails.Add((Path.GetFileName(filePath), "Données email non extraites (structure incorrecte ou vide)."));
                continue;
            }

            var partnerNameFromFile = emailData.PartnerNameInFile?.NormalizeSpaces() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(partnerNameFromFile))
            {
                await LogAndSend(
                    $"***************************************************************************************************************************\n" +
                    $"Aucun nom de partenaire extrait dans le fichier '{Path.GetFileName(filePath)}'. Impossible de trouver une correspondance.",
                    cancellationToken);

                ignoredFilesDetails.Add((Path.GetFileName(filePath), "Nom de partenaire non détecté dans le fichier."));
                continue;
            }

            var partner = partners.FirstOrDefault(p =>
                p.PartnerName.NormalizeSpaces().Equals(partnerNameFromFile, StringComparison.OrdinalIgnoreCase));

            if (partner != null && partner.Emails.Any())
            {
                preparedEmails.Add(new EmailToSend
                {
                    Subject = _emailContentBuilder.BuildSubject(emailData),
                    Body = _emailContentBuilder.BuildBody(emailData),
                    ToRecipients = partner.Emails,
                    CcRecipients = ccRecipients ?? new(),
                    BccRecipients = bccRecipients ?? new(),
                    FromDisplayName = fromDisplayName ?? string.Empty,
                    AttachmentFilePaths = new List<string> { filePath },
                    SmtpHost = smtpHost,
                    SmtpPort = smtpPort,
                    SmtpFromEmail = smtpFromEmail,
                    PartnerName = partner.PartnerName
                });
            }
            else
            {
                var closestPartner = partners
                    .Select(p => new
                    {
                        Partner = p,
                        Distance = StringUtils.ComputeLevenshteinDistance(
                            p.PartnerName.NormalizeSpaces().ToLowerInvariant(),
                            partnerNameFromFile.ToLowerInvariant())
                    })
                    .OrderBy(x => x.Distance)
                    .FirstOrDefault();

                string closestName = closestPartner?.Partner.PartnerName ?? "[AUCUN PARTENAIRE]";
                int distance = closestPartner?.Distance ?? -1;

                string diffDetails = StringUtils.ShowStringDifferences(closestName, partnerNameFromFile);

                await LogAndSend(
                    $"***************************************************************************************************************************\n" +
                    $"Aucun partenaire exact trouvé pour le fichier '{Path.GetFileName(filePath)}' (Nom extrait: '{partnerNameFromFile}').\n" +
                    $"Nom partenaire le plus proche: '{closestName}' (Distance: {distance}).\n" +
                    $"Détail des différences :\n{diffDetails}",
                    cancellationToken);

                ignoredFilesDetails.Add((Path.GetFileName(filePath), $"Aucun partenaire ne correspond au nom extrait '{partnerNameFromFile}' — plus proche : '{closestName}' (distance : {distance})"));
            }
        }

        // 🔚 Log final de récapitulatif des fichiers ignorés
        if (ignoredFilesDetails.Any())
        {
            await LogAndSend("\n\n📋 Résumé final des fichiers ignorés :", cancellationToken);
            foreach (var entry in ignoredFilesDetails)
            {
                await LogAndSend($"❌ {entry.FileName} — {entry.Reason}", cancellationToken);
            }
            await LogAndSend($"➡️ Total ignorés : {ignoredFilesDetails.Count} fichier(s)", cancellationToken);
        }
        else
        {
            await LogAndSend("\n✅ Tous les fichiers ont été traités avec succès (aucun fichier ignoré).", cancellationToken);
        }

        return preparedEmails;
    }



















    public async Task<List<EmailSendResult>> SendPreparedEmailsAsync(
     List<EmailToSend> emails,
     CancellationToken cancellationToken)
    {
        var results = new List<EmailSendResult>();

        foreach (var email in emails)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                bool sent = await _sendEmailService.SendEmailAsync(
                    subject: email.Subject,
                    body: email.Body,
                    toRecipients: email.ToRecipients,
                    ccRecipients: email.CcRecipients,
                    bccRecipients: email.BccRecipients,
                    fromDisplayName: email.FromDisplayName,
                    attachmentFilePaths: email.AttachmentFilePaths,
                    smtpHost: email.SmtpHost,
                    smtpPort: email.SmtpPort,
                    smtpFromEmail: email.SmtpFromEmail,
                    cancellationToken: cancellationToken
                );

                if (sent)
                {

                    results.Add(new EmailSendResult
                    {
                        To = string.Join(", ", email.ToRecipients),
                        Success = true,
                        ErrorMessage = null
                    });
                }
                else
                {
                    await LogAndSendError($"Échec de l'envoi de l'email à : {string.Join(", ", email.ToRecipients)}", cancellationToken);

                    results.Add(new EmailSendResult
                    {
                        To = string.Join(", ", email.ToRecipients),
                        Success = false,
                        ErrorMessage = "Échec inconnu de l'envoi"
                    });
                }
            }
            catch (Exception ex)
            {
                results.Add(new EmailSendResult
                {
                    To = string.Join(", ", email.ToRecipients),
                    Success = false,
                    ErrorMessage = ex.Message
                });
            }
        }

        return results;
    }

}


