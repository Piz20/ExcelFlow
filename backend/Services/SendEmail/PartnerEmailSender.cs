// Fichier : Services/PartnerEmailSender.cs
using MailKit.Net.Smtp; // If still needed for other SendEmail dependency logic, otherwise can remove
using MailKit.Security; // If still needed for other SendEmail dependency logic, otherwise can remove
using MimeKit; // If still needed for other SendEmail dependency logic, otherwise can remove
using Microsoft.Extensions.Configuration; // If still needed for other SendEmail dependency logic, otherwise can remove
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
using ExcelFlow.Models; // <<< ADDED for PartnerInfo, SentEmailSummary, and EmailData

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

    // PartnerInfo and SentEmailSummary are now in ExcelFlow.Models

    public async Task SendEmailsToPartnersWithAttachments(
        string partnerEmailFilePath,
        string generatedFilesFolderPath,
        string? fromDisplayName = null,
        List<string>? ccRecipients = null,
        List<string>? bccRecipients = null,
        CancellationToken cancellationToken = default)
    {
        await LogAndSend($"Démarrage du processus d'envoi d'emails basé sur les fichiers générés...", cancellationToken);

        await LogAndSend($"Lecture des partenaires depuis : {partnerEmailFilePath}...", cancellationToken);
        List<PartnerInfo> partners;
        try
        {
            partners = await Task.Run(() => _partnerExcelReader.ReadPartnersFromExcel(partnerEmailFilePath, cancellationToken), cancellationToken);
            await LogAndSend($"Lecture terminée. {partners.Count} partenaires trouvés dans le fichier Excel.", cancellationToken);

            if (partners.Any())
            {
                await LogAndSend("Détails des partenaires lus (Nom et Adresses Email) :", cancellationToken);
                foreach (var p in partners)
                {
                    string searchableNames = $"'{p.SearchableNameFull}'";
                    if (p.SearchableNameSigle != null)
                    {
                        searchableNames += $" (Sigle: '{p.SearchableNameSigle}')";
                    }
                    await LogAndSend($"   - Partenaire: '{p.PartnerName}' | Emails: {string.Join(", ", p.Emails)} | Noms de recherche normalisés: {searchableNames}", cancellationToken);
                }
            }
            else
            {
                await LogAndSend("Aucun partenaire avec adresse email valide n'a été trouvé pour le logging initial.", cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            await LogAndSendError("L'opération de lecture du fichier Excel a été annulée.", cancellationToken);
            return;
        }
        catch (Exception ex)
        {
            await LogAndSendError($"Erreur lors de la lecture du fichier Excel : {ex.Message}", cancellationToken);
            return;
        }

        if (!partners.Any())
        {
            await LogAndSend("Aucun partenaire trouvé dans le fichier Excel ou aucune adresse email valide. Aucun email ne sera envoyé.", cancellationToken);
            return;
        }

        await LogAndSend($"Analyse du dossier des fichiers générés : {generatedFilesFolderPath}...", cancellationToken);
        if (!Directory.Exists(generatedFilesFolderPath))
        {
            await LogAndSendError($"Le dossier des fichiers générés est introuvable : {generatedFilesFolderPath}. Aucun email avec pièce jointe ne sera envoyé.", cancellationToken);
            return;
        }

        var allGeneratedFiles = Directory.GetFiles(generatedFilesFolderPath, "*", SearchOption.TopDirectoryOnly).ToList();
        await LogAndSend($"Trouvé {allGeneratedFiles.Count} fichiers potentiels à envoyer.", cancellationToken);

        if (!allGeneratedFiles.Any())
        {
            await LogAndSend("Aucun fichier généré trouvé dans le dossier. Aucun email ne sera envoyé.", cancellationToken);
            return;
        }

        var processedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sentEmailSummaries = new List<SentEmailSummary>();

        int filesProcessedCount = 0;
        foreach (var filePath in allGeneratedFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (processedFiles.Contains(filePath))
            {
                await LogAndSend($"   Fichier '{Path.GetFileName(filePath)}' déjà traité pour un autre partenaire. Ignoré.", cancellationToken);
                await LogAndSend("---", cancellationToken);
                continue;
            }

            filesProcessedCount++;
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            await LogAndSend($"Traitement du fichier {filesProcessedCount}/{allGeneratedFiles.Count}: '{Path.GetFileName(filePath)}'", cancellationToken);

            string normalizedFileName = _partnerExcelReader.NormalizeForComparison(fileNameWithoutExtension); // Using method from PartnerExcelReader

            PartnerInfo? foundPartner = null;

            foreach (var partner in partners)
            {
                cancellationToken.ThrowIfCancellationRequested();

                bool matched = false;

                if (!string.IsNullOrWhiteSpace(partner.SearchableNameFull))
                {
                    var regexFull = new Regex($@"\b{Regex.Escape(partner.SearchableNameFull)}\b", RegexOptions.CultureInvariant);
                    if (regexFull.IsMatch(normalizedFileName))
                    {
                        matched = true;
                    }
                }

                if (!matched && partner.SearchableNameSigle != null && !string.IsNullOrWhiteSpace(partner.SearchableNameSigle))
                {
                    var regexSigle = new Regex($@"\b{Regex.Escape(partner.SearchableNameSigle)}\b", RegexOptions.CultureInvariant);
                    if (regexSigle.IsMatch(normalizedFileName))
                    {
                        matched = true;
                    }
                }

                if (matched)
                {
                    foundPartner = partner;
                    break;
                }
            }

            if (foundPartner != null && foundPartner.Emails.Any())
            {
                await LogAndSend($"   Fichier '{Path.GetFileName(filePath)}' contient le nom/sigle exact et ordonné du partenaire '{foundPartner.PartnerName}'.", cancellationToken);
                await LogAndSend($"   Adresses email cibles pour '{foundPartner.PartnerName}': {string.Join(", ", foundPartner.Emails)}", cancellationToken);

                // Now EmailData is from ExcelFlow.Models
                var emailData = await _emailDataExtractor.ExtractEmailDataFromAttachment(filePath, cancellationToken);
                if (emailData == null)
                {
                    await LogAndSendError($"Impossible d'extraire les données dynamiques pour l'email du fichier '{Path.GetFileName(filePath)}'. Email non envoyé.", cancellationToken);
                    await LogAndSend("---", cancellationToken);
                    continue;
                }

                string finalSubject = _emailContentBuilder.BuildSubject(emailData);
                string finalBody = _emailContentBuilder.BuildBody(emailData);

                await LogAndSend($"   Envoi de l'email à {foundPartner.PartnerName} avec le fichier '{Path.GetFileName(filePath)}' en pièce jointe.", cancellationToken);

                bool sent = await _sendEmailService.SendEmailAsync(
                    subject: finalSubject,
                    body: finalBody,
                    toRecipients: foundPartner.Emails,
                    ccRecipients: ccRecipients,
                    bccRecipients: bccRecipients,
                    fromDisplayName: fromDisplayName,
                    attachmentFilePaths: new List<string> { filePath },
                    cancellationToken: cancellationToken
                );

                if (sent)
                {
                    await LogAndSend($"Email envoyé avec succès à {foundPartner.PartnerName} pour le fichier '{Path.GetFileName(filePath)}'.", cancellationToken);
                    processedFiles.Add(filePath);
                    sentEmailSummaries.Add(new SentEmailSummary
                    {
                        FileName = Path.GetFileName(filePath),
                        PartnerName = foundPartner.PartnerName,
                        RecipientEmails = foundPartner.Emails.ToList(),
                        CcRecipientsSent = ccRecipients?.ToList() ?? new List<string>(),
                        BccRecipientsSent = bccRecipients?.ToList() ?? new List<string>()
                    });
                }
                else
                {
                    await LogAndSendError($"Échec de l'envoi de l'email à {foundPartner.PartnerName} pour le fichier '{Path.GetFileName(filePath)}'.", cancellationToken);
                }
            }
            else
            {
                if (foundPartner != null && !foundPartner.Emails.Any())
                {
                    await LogAndSendError($"   Partenaire '{foundPartner.PartnerName}' trouvé pour le fichier '{Path.GetFileName(filePath)}' mais sans adresse email valide. Email non envoyé.", cancellationToken);
                }
                else
                {
                    await LogAndSend($"   Aucun partenaire dont le nom/sigle exact et ordonné est inclus comme un mot entier dans le fichier '{Path.GetFileName(filePath)}'. Le fichier sera ignoré.", cancellationToken);
                }
            }
            await LogAndSend("---", cancellationToken);
        }
        await LogAndSend("Processus d'envoi d'emails basé sur les fichiers générés terminé.", cancellationToken);

        await LogAndSend("\n--- RÉCAPITULATIF DES EMAILS ENVOYÉS ---", cancellationToken);
        if (sentEmailSummaries.Any())
        {
            await LogAndSend($"Total de fichiers envoyés : {sentEmailSummaries.Count}", cancellationToken);
            foreach (var summary in sentEmailSummaries)
            {
                await LogAndSend($"   - Fichier: '{summary.FileName}' envoyé à Partenaire: '{summary.PartnerName}' (To: {string.Join(", ", summary.RecipientEmails)})", cancellationToken);
                if (summary.CcRecipientsSent.Any())
                {
                    await LogAndSend($"     Cc: {string.Join(", ", summary.CcRecipientsSent)}", cancellationToken);
                }
                if (summary.BccRecipientsSent.Any())
                {
                    await LogAndSend($"     Bcc: {string.Join(", ", summary.BccRecipientsSent)} (Non visible par les destinataires To/Cc)", cancellationToken);
                }
            }
        }
        else
        {
            await LogAndSend("Aucun email n'a été envoyé durant ce processus.", cancellationToken);
        }
        await LogAndSend("---------------------------------------", cancellationToken);
    }
}