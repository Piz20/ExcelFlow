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

    // Helper pour envoyer des mises à jour de progression structurées directement
    // et EN PLUS, un log formaté du pourcentage.
    private async Task SendProgressToFrontend(int current, int total, string message, CancellationToken cancellationToken = default)
    {
        var progress = new ProgressUpdate
        {
            Current = current,
            Total = total,
            Percentage = total > 0 ? (int)((double)current / total * 100) : 0,
            Message = message
        };

        // 1. Envoyer l'objet ProgressUpdate structuré pour la barre de progression/UI
        if (_hubContext != null)
        {
            await _hubContext.Clients.All.SendAsync("ReceiveProgressUpdate", progress, cancellationToken);
        }

        // 2. Envoyer une ligne de log séparée avec le pourcentage formaté
        string percentageLine = $"----------------------------------------------{progress.Percentage}%";
        await LogAndSend(percentageLine, cancellationToken); // Utilise ReceiveMessage

        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] PROGRESSION: {message} ({progress.Percentage}%)");
    }

    // Helper pour envoyer la liste des partenaires identifiés
    private async Task SendIdentifiedPartnersToFrontend(List<PartnerInfo> partners, CancellationToken cancellationToken = default)
    {
        if (_hubContext != null)
        {
            await _hubContext.Clients.All.SendAsync("ReceiveIdentifiedPartners", partners, cancellationToken);
        }
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] PARTENAIRES IDENTIFIÉS : {partners.Count} partenaires envoyés au frontend.");
    }

    // Helper pour envoyer un récapitulatif d'email envoyé
    private async Task SendSentEmailSummaryToFrontend(SentEmailSummary summary, CancellationToken cancellationToken = default)
    {
        if (_hubContext != null)
        {
            await _hubContext.Clients.All.SendAsync("ReceiveSentEmailSummary", summary, cancellationToken);
        }
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] RÉCAPITULATIF EMAIL : Fichier '{summary.FileName}' envoyé à '{summary.PartnerName}'.");
    }

    // Helper pour envoyer le nombre total de fichiers à traiter
    private async Task SendTotalFilesCountToFrontend(int totalFiles, CancellationToken cancellationToken = default)
    {
        if (_hubContext != null)
        {
            await _hubContext.Clients.All.SendAsync("ReceiveTotalFilesCount", totalFiles, cancellationToken);
        }
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] TOTAL FICHIERS : {totalFiles} fichiers détectés pour l'envoi.");
    }








    public static int ComputeLevenshteinDistance(string s, string t)
    {
        int n = s.Length;
        int m = t.Length;
        var d = new int[n + 1, m + 1];

        for (int i = 0; i <= n; i++) d[i, 0] = i;
        for (int j = 0; j <= m; j++) d[0, j] = j;

        for (int i = 1; i <= n; i++)
        {
            for (int j = 1; j <= m; j++)
            {
                int cost = (s[i - 1] == t[j - 1]) ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }
        }
        return d[n, m];
    }

    public static string ShowStringDifferences(string original, string compared)
    {
        int maxLength = Math.Max(original.Length, compared.Length);
        var diffBuilder = new System.Text.StringBuilder();

        diffBuilder.AppendLine($"Original : \"{original}\"");
        diffBuilder.AppendLine($"Compared : \"{compared}\"");
        diffBuilder.AppendLine("Diff    : ");

        for (int i = 0; i < maxLength; i++)
        {
            char c1 = i < original.Length ? original[i] : '-';
            char c2 = i < compared.Length ? compared[i] : '-';

            diffBuilder.Append(c1 == c2 ? ' ' : '^');
        }

        return diffBuilder.ToString();
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

        for (int i = 0; i < allGeneratedFiles.Count; i++)
        {
            var filePath = allGeneratedFiles[i];

            await LogAndSend(
      $"\n\n\n\n📄 Traitement du fichier {i + 1}/{allGeneratedFiles.Count} : {Path.GetFileName(filePath)}",
      cancellationToken);

            var emailData = await _emailDataExtractor.ExtractEmailDataFromAttachment(filePath, cancellationToken);
            if (emailData == null)
            {
                continue;
            }

            var partnerNameFromFile = emailData.PartnerNameInFile?.NormalizeSpaces() ?? string.Empty;

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
                if (!string.IsNullOrEmpty(partnerNameFromFile))
                {
                    var closestPartner = partners
                        .Select(p => new
                        {
                            Partner = p,
                            Distance = ComputeLevenshteinDistance(
                                p.PartnerName.NormalizeSpaces().ToLowerInvariant(),
                                partnerNameFromFile.ToLowerInvariant())
                        })
                        .OrderBy(x => x.Distance)
                        .FirstOrDefault();

                    string closestName = closestPartner?.Partner.PartnerName ?? "[AUCUN PARTENAIRE]";
                    int distance = closestPartner?.Distance ?? -1;

                    string diffDetails = ShowStringDifferences(closestName, partnerNameFromFile);

                    await LogAndSend(
                        $"*****************************************************************************************************************************************************************************************************************************\n" +
                        $"Aucun partenaire exact trouvé pour le fichier '{Path.GetFileName(filePath)}' (Nom extrait: '{partnerNameFromFile}').\n" +
                        $"Nom partenaire le plus proche: '{closestName}' (Distance: {distance}).\n" +
                        $"Détail des différences :\n{diffDetails}",
                        cancellationToken);
                }
                else
                {
                    await LogAndSend(
                        $"*****************************************************************************************************************************************************************************************************************************\n" +
                        $"Aucun nom de partenaire extrait dans le fichier '{Path.GetFileName(filePath)}'. Impossible de trouver une correspondance.",
                        cancellationToken);
                }
            }
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
                    await LogAndSend($"Email envoyé à : {string.Join(", ", email.ToRecipients)}", cancellationToken);

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


