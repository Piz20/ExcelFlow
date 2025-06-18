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

public class PartnerEmailSender
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


public async Task SendEmailsToPartnersWithAttachments(
    string partnerEmailFilePath,
    string generatedFilesFolderPath,
    string? smtpFromEmail = null,
    string? smtpHost = null,
    int? smtpPort = null,
    List<string>? ccRecipients = null,
    List<string>? bccRecipients = null,
    CancellationToken cancellationToken = default)
{
    await LogAndSend("Démarrage du processus d'envoi d'emails basé sur les fichiers générés...", cancellationToken);
    await SendProgressToFrontend(0, 0, "Démarrage de l'opération.", cancellationToken);

    // Lecture partenaires
    List<PartnerInfo> partners;
    try
    {
        partners = await Task.Run(() => _partnerExcelReader.ReadPartnersFromExcel(partnerEmailFilePath, cancellationToken), cancellationToken);
        await LogAndSend($"Lecture terminée. {partners.Count} partenaires trouvés.", cancellationToken);
        await SendProgressToFrontend(0, 0, $"Lecture des partenaires terminée. {partners.Count} partenaires trouvés.", cancellationToken);

        if (!partners.Any())
        {
            await LogAndSend("Aucun partenaire avec adresse email valide trouvé. Arrêt de l'opération.", cancellationToken);
            await SendProgressToFrontend(0, 0, "Aucun partenaire trouvé. Arrêt.", cancellationToken);
            return;
        }
    }
    catch (OperationCanceledException)
    {
        await LogAndSendError("Lecture du fichier Excel annulée.", cancellationToken);
        await SendProgressToFrontend(0, 0, "Opération annulée pendant la lecture des partenaires.", cancellationToken);
        return;
    }
    catch (Exception ex)
    {
        await LogAndSendError($"Erreur lecture fichier Excel : {ex.Message}", cancellationToken);
        await SendProgressToFrontend(0, 0, $"Erreur lors de la lecture des partenaires : {ex.Message}", cancellationToken);
        return;
    }

    if (!Directory.Exists(generatedFilesFolderPath))
    {
        await LogAndSendError($"Dossier des fichiers générés introuvable : {generatedFilesFolderPath}", cancellationToken);
        await SendProgressToFrontend(0, 0, "Erreur : dossier des fichiers générés introuvable.", cancellationToken);
        return;
    }

    var allGeneratedFiles = Directory.GetFiles(generatedFilesFolderPath).ToList();
    if (!allGeneratedFiles.Any())
    {
        await LogAndSend("Aucun fichier généré trouvé. Aucun email ne sera envoyé.", cancellationToken);
        await SendProgressToFrontend(0, 0, "Processus terminé: Aucun fichier à envoyer.", cancellationToken);
        return;
    }

    await LogAndSend($"Traitement de {allGeneratedFiles.Count} fichiers.", cancellationToken);
    await SendTotalFilesCountToFrontend(allGeneratedFiles.Count, cancellationToken);

    var processedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var sentEmailSummaries = new List<SentEmailSummary>();

    int filesProcessedCount = 0;
    int emailsSentSuccessfully = 0;

    foreach (var filePath in allGeneratedFiles)
    {
        cancellationToken.ThrowIfCancellationRequested();

        string fileName = Path.GetFileName(filePath);

        if (processedFiles.Contains(filePath))
        {
            filesProcessedCount++;
            await LogAndSend($"[Fichier '{fileName}'] Déjà traité, ignoré.", cancellationToken);
            await SendProgressToFrontend(filesProcessedCount, allGeneratedFiles.Count, $"'{fileName}' déjà traité.", cancellationToken);
            continue;
        }

        filesProcessedCount++;

        var emailData = await _emailDataExtractor.ExtractEmailDataFromAttachment(filePath, cancellationToken);
        if (emailData == null)
        {
            await LogAndSendError($"[Fichier '{fileName}'] Impossible d'extraire les données dynamiques. Email non envoyé.", cancellationToken);
            await SendProgressToFrontend(filesProcessedCount, allGeneratedFiles.Count, $"'{fileName}': Extraction données échouée.", cancellationToken);
            continue;
        }
        var partnerNameInFile = (emailData.PartnerNameInFile ?? string.Empty).Trim();

        PartnerInfo? foundPartner = null;

        foreach (var partner in partners)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (partner.PartnerName.Trim().Equals(partnerNameInFile, StringComparison.OrdinalIgnoreCase))
            {
                foundPartner = partner;
                break;
            }
        }

        if (foundPartner != null && foundPartner.Emails.Any())
        {
            await LogAndSend($"[Fichier '{fileName}'] Partenaire '{foundPartner.PartnerName}' trouvé.", cancellationToken);

            string finalSubject = _emailContentBuilder.BuildSubject(emailData);
            string finalBody = _emailContentBuilder.BuildBody(emailData);

            bool sent = await _sendEmailService.SendEmailAsync(
                subject: finalSubject,
                body: finalBody,
                toRecipients: foundPartner.Emails,
                ccRecipients: ccRecipients,
                bccRecipients: bccRecipients,
                fromDisplayName: smtpFromEmail,
                attachmentFilePaths: new List<string> { filePath },
                smtpHost: smtpHost,
                smtpPort: smtpPort,
                smtpFromEmail: smtpFromEmail ,
                cancellationToken: cancellationToken
            );

            if (sent)
            {
                emailsSentSuccessfully++;
                processedFiles.Add(filePath);

                var summary = new SentEmailSummary
                {
                    FileName = fileName,
                    PartnerName = foundPartner.PartnerName,
                    RecipientEmails = foundPartner.Emails.ToList(),
                    CcRecipientsSent = ccRecipients?.ToList() ?? new List<string>(),
                    BccRecipientsSent = bccRecipients?.ToList() ?? new List<string>()
                };
                sentEmailSummaries.Add(summary);

                await SendSentEmailSummaryToFrontend(summary, cancellationToken);
                await SendProgressToFrontend(filesProcessedCount, allGeneratedFiles.Count, $"'{fileName}': Email envoyé avec succès.", cancellationToken);
            }
            else
            {
                await LogAndSendError($"[Fichier '{fileName}'] Échec de l'envoi à {foundPartner.PartnerName}.", cancellationToken);
                await SendProgressToFrontend(filesProcessedCount, allGeneratedFiles.Count, $"'{fileName}': Échec de l'envoi de l'email.", cancellationToken);
            }
        }
        else
        {
            if (foundPartner != null && !foundPartner.Emails.Any())
            {
                await LogAndSendError($"[Fichier '{fileName}'] Partenaire '{foundPartner.PartnerName}' sans email valide. Fichier ignoré.", cancellationToken);
                await SendProgressToFrontend(filesProcessedCount, allGeneratedFiles.Count, $"'{fileName}': Partenaire sans email valide.", cancellationToken);
            }
            else
            {
                await LogAndSend($"[Fichier '{fileName}'] Aucun partenaire correspondant trouvé. Fichier ignoré.", cancellationToken);
                await SendProgressToFrontend(filesProcessedCount, allGeneratedFiles.Count, $"'{fileName}': Aucun partenaire correspondant.", cancellationToken);
            }
        }
    }

    await LogAndSend($"Processus terminé. Total d'emails envoyés : {emailsSentSuccessfully}.", cancellationToken);
    await SendProgressToFrontend(allGeneratedFiles.Count, allGeneratedFiles.Count, $"Processus terminé. {emailsSentSuccessfully} emails envoyés.", cancellationToken);

    if (sentEmailSummaries.Any())
    {
        await LogAndSend("--- RÉCAPITULATIF FINAL ---", cancellationToken);
        await LogAndSend($"Total d'emails envoyés avec succès : {sentEmailSummaries.Count}", cancellationToken);
        foreach (var summary in sentEmailSummaries)
        {
            await LogAndSend($"- '{summary.FileName}' envoyé à '{summary.PartnerName}' (To: {string.Join(", ", summary.RecipientEmails)})", cancellationToken);
            if (summary.CcRecipientsSent.Any())
                await LogAndSend($"  Cc: {string.Join(", ", summary.CcRecipientsSent)}", cancellationToken);
            if (summary.BccRecipientsSent.Any())
                await LogAndSend($"  Bcc: {string.Join(", ", summary.BccRecipientsSent)} (Non visible par To/Cc)", cancellationToken);
        }
        await LogAndSend("---------------------------------------", cancellationToken);
    }
    else
    {
        await LogAndSend("Aucun email envoyé durant ce processus.", cancellationToken);
    }
}


}