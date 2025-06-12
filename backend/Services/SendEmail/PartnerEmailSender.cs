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
        string? fromDisplayName = null,
        List<string>? ccRecipients = null,
        List<string>? bccRecipients = null,
        CancellationToken cancellationToken = default)
    {
        // --- 1. Démarrage du processus ---
        await LogAndSend($"Démarrage du processus d'envoi d'emails basé sur les fichiers générés...", cancellationToken);
        await SendProgressToFrontend(0, 0, "Démarrage de l'opération.", cancellationToken);

        // --- 2. Lecture des partenaires ---
        await LogAndSend($"Lecture des partenaires depuis : {partnerEmailFilePath}...", cancellationToken);
        List<PartnerInfo> partners;
        try
        {
            // Ligne corrigée ici : Appel à ReadPartnersFromExcel qui retourne une List<PartnerInfo>
            partners = await Task.Run(() => _partnerExcelReader.ReadPartnersFromExcel(partnerEmailFilePath, cancellationToken), cancellationToken);
            await LogAndSend($"Lecture terminée. {partners.Count} partenaires trouvés dans le fichier Excel.", cancellationToken);
            await SendProgressToFrontend(0, 0, $"Lecture des partenaires terminée. {partners.Count} partenaires trouvés.", cancellationToken);

            if (partners.Any())
            {
                await SendIdentifiedPartnersToFrontend(partners, cancellationToken);
                await LogAndSend("Détails des partenaires lus (Nom et Adresses Email) :", cancellationToken);
                foreach (var p in partners)
                {
                    string searchableNames = $"'{p.SearchableNameFull}'";
                    if (p.SearchableNameSigle != null)
                    {
                        searchableNames += $" (Sigle: '{p.SearchableNameSigle}')";
                    }
                    await LogAndSend($"    - Partenaire: '{p.PartnerName}' | Emails: {string.Join(", ", p.Emails)} | Noms de recherche normalisés: {searchableNames}", cancellationToken);
                }
            }
            else
            {
                await LogAndSend("Aucun partenaire avec adresse email valide n'a été trouvé.", cancellationToken);
                await SendProgressToFrontend(0, 0, "Aucun partenaire trouvé. Arrêt de l'opération.", cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            await LogAndSendError("L'opération de lecture du fichier Excel a été annulée.", cancellationToken);
            await SendProgressToFrontend(0, 0, "Opération annulée pendant la lecture des partenaires.", cancellationToken);
            return;
        }
        catch (Exception ex)
        {
            await LogAndSendError($"Erreur lors de la lecture du fichier Excel : {ex.Message}", cancellationToken);
            await SendProgressToFrontend(0, 0, $"Erreur lors de la lecture des partenaires : {ex.Message}", cancellationToken);
            return;
        }

        if (!partners.Any())
        {
            await LogAndSend("Aucun partenaire trouvé dans le fichier Excel ou aucune adresse email valide. Aucun email ne sera envoyé.", cancellationToken);
            await SendProgressToFrontend(0, 0, "Processus terminé: Aucun partenaire pour l'envoi d'emails.", cancellationToken);
            return;
        }

        // --- 3. Analyse du dossier des fichiers générés ---
        await LogAndSend($"Analyse du dossier des fichiers générés : {generatedFilesFolderPath}...", cancellationToken);
        await SendProgressToFrontend(0, 0, "Analyse du dossier des fichiers générés...", cancellationToken);

        if (!Directory.Exists(generatedFilesFolderPath))
        {
            await LogAndSendError($"Le dossier des fichiers générés est introuvable : {generatedFilesFolderPath}. Aucun email avec pièce jointe ne sera envoyé.", cancellationToken);
            await SendProgressToFrontend(0, 0, $"Erreur: Dossier des fichiers générés introuvable.", cancellationToken);
            return;
        }

        var allGeneratedFiles = Directory.GetFiles(generatedFilesFolderPath, "*", SearchOption.TopDirectoryOnly).ToList();
        
        int totalFilesToProcess = allGeneratedFiles.Count;
        await LogAndSend($"Trouvé {totalFilesToProcess} fichiers potentiels à envoyer.", cancellationToken);
        
        await SendTotalFilesCountToFrontend(totalFilesToProcess, cancellationToken);
        await SendProgressToFrontend(0, totalFilesToProcess, $"Début du traitement de {totalFilesToProcess} fichiers.", cancellationToken);


        if (!allGeneratedFiles.Any())
        {
            await LogAndSend("Aucun fichier généré trouvé dans le dossier. Aucun email ne sera envoyé.", cancellationToken);
            await SendProgressToFrontend(0, 0, "Processus terminé: Aucun fichier à envoyer.", cancellationToken);
            return;
        }

        var processedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sentEmailSummaries = new List<SentEmailSummary>();

        int filesProcessedCount = 0;
        int emailsSentSuccessfully = 0;

        // --- 4. Parcours des fichiers et envoi des emails ---
        foreach (var filePath in allGeneratedFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();

            string fileName = Path.GetFileName(filePath);
            
            // La logique 'processedFiles.Contains(filePath)' est un peu redondante si on ne traite
            // qu'un seul fichier par partenaire. Si un fichier ne doit être envoyé qu'une fois,
            // même s'il correspond à plusieurs mots-clés de différents partenaires, cette logique est utile.
            // Sinon, si chaque fichier correspond à un seul partenaire, cette partie peut être simplifiée.
            // Pour l'instant, je la garde pour éviter de modifier la logique de déduplication existante.
            if (processedFiles.Contains(filePath))
            {
                filesProcessedCount++;
                await LogAndSend($"[Fichier '{fileName}'] Fichier déjà traité. Ignoré.", cancellationToken);
                await SendProgressToFrontend(filesProcessedCount, totalFilesToProcess, $"Fichier '{fileName}' déjà traité. Ignoré.", cancellationToken);
                await LogAndSend("---", cancellationToken);
                continue;
            }

            filesProcessedCount++; // Incrément pour le fichier en cours de traitement
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            
            await LogAndSend($"[Fichier '{fileName}'] Traitement en cours ({filesProcessedCount}/{totalFilesToProcess})...", cancellationToken);
            await SendProgressToFrontend(filesProcessedCount -1, totalFilesToProcess, $"Traitement de '{fileName}'. Recherche du partenaire.", cancellationToken);

            string normalizedFileName = _partnerExcelReader.NormalizeForComparison(fileNameWithoutExtension);

            PartnerInfo? foundPartner = null;

            foreach (var partner in partners)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // On vérifie tous les mots-clés de recherche du partenaire
                foreach (var keyword in partner.SearchableKeywords)
                {
                    // Utilisation de Regex.Escape pour sécuriser le mot-clé dans l'expression régulière
                    // et \b pour les limites de mot afin d'éviter les correspondances partielles indésirables.
                    var regex = new Regex($@"\b{Regex.Escape(keyword)}\b", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
                    if (regex.IsMatch(normalizedFileName))
                    {
                        foundPartner = partner;
                        break; // Partenaire trouvé pour ce fichier, pas besoin de vérifier d'autres mots-clés pour ce fichier/partenaire
                    }
                }
                if (foundPartner != null)
                {
                    break; // Partenaire trouvé pour ce fichier, pas besoin de vérifier d'autres partenaires
                }
            }

            if (foundPartner != null && foundPartner.Emails.Any())
            {
                await LogAndSend($"[Fichier '{fileName}'] Partenaire '{foundPartner.PartnerName}' trouvé. Adresses email: {string.Join(", ", foundPartner.Emails)}.", cancellationToken);
                await SendProgressToFrontend(filesProcessedCount -1, totalFilesToProcess, $"'{fileName}': Partenaire '{foundPartner.PartnerName}' trouvé. Extraction des données.", cancellationToken);

                var emailData = await _emailDataExtractor.ExtractEmailDataFromAttachment(filePath, cancellationToken);
                if (emailData == null)
                {
                    await LogAndSendError($"[Fichier '{fileName}'] Impossible d'extraire les données dynamiques pour l'email. Email non envoyé.", cancellationToken);
                    await SendProgressToFrontend(filesProcessedCount, totalFilesToProcess, $"'{fileName}': Échec d'extraction des données. Email non envoyé.", cancellationToken);
                    await LogAndSend("---", cancellationToken);
                    continue;
                }

                string finalSubject = _emailContentBuilder.BuildSubject(emailData);
                string finalBody = _emailContentBuilder.BuildBody(emailData);

                await LogAndSend($"[Fichier '{fileName}'] Envoi de l'email à {foundPartner.PartnerName}...", cancellationToken);
                await SendProgressToFrontend(filesProcessedCount -1, totalFilesToProcess, $"'{fileName}': Tentative d'envoi à {foundPartner.PartnerName}.", cancellationToken);

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
                    emailsSentSuccessfully++;
                    await LogAndSend($"[Fichier '{fileName}'] Email envoyé avec succès à {foundPartner.PartnerName}.", cancellationToken);
                    processedFiles.Add(filePath); // Marquer le fichier comme traité
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
                    await SendProgressToFrontend(filesProcessedCount, totalFilesToProcess, $"'{fileName}': Email envoyé avec succès.", cancellationToken);
                }
                else
                {
                    await LogAndSendError($"[Fichier '{fileName}'] Échec de l'envoi de l'email à {foundPartner.PartnerName}.", cancellationToken);
                    await SendProgressToFrontend(filesProcessedCount, totalFilesToProcess, $"'{fileName}': Échec de l'envoi de l'email.", cancellationToken);
                }
            }
            else
            {
                if (foundPartner != null && !foundPartner.Emails.Any())
                {
                    await LogAndSendError($"[Fichier '{fileName}'] Partenaire '{foundPartner.PartnerName}' trouvé mais sans adresse email valide. Fichier ignoré.", cancellationToken);
                    await SendProgressToFrontend(filesProcessedCount, totalFilesToProcess, $"'{fileName}': Partenaire sans email valide. Fichier ignoré.", cancellationToken);
                }
                else
                {
                    await LogAndSend($"[Fichier '{fileName}'] Aucun partenaire correspondant trouvé. Fichier ignoré.", cancellationToken);
                    await SendProgressToFrontend(filesProcessedCount, totalFilesToProcess, $"'{fileName}': Aucun partenaire correspondant. Fichier ignoré.", cancellationToken);
                }
            }
            await LogAndSend("---", cancellationToken); // Séparateur pour la lisibilité
        }

        await LogAndSend("Processus d'envoi d'emails basé sur les fichiers générés terminé.", cancellationToken);
        await SendProgressToFrontend(totalFilesToProcess, totalFilesToProcess, $"Processus terminé. Total d'emails envoyés : {emailsSentSuccessfully}.", cancellationToken);

        await LogAndSend("\n--- RÉCAPITULATIF FINAL DES EMAILS ---", cancellationToken);
        if (sentEmailSummaries.Any())
        {
            await LogAndSend($"Total d'emails envoyés avec succès : {sentEmailSummaries.Count}", cancellationToken);
            foreach (var summary in sentEmailSummaries)
            {
                await LogAndSend($"    - Fichier: '{summary.FileName}' envoyé à Partenaire: '{summary.PartnerName}' (To: {string.Join(", ", summary.RecipientEmails)})", cancellationToken);
                if (summary.CcRecipientsSent.Any())
                {
                    await LogAndSend($"      Cc: {string.Join(", ", summary.CcRecipientsSent)}", cancellationToken);
                }
                if (summary.BccRecipientsSent.Any())
                {
                    await LogAndSend($"      Bcc: {string.Join(", ", summary.BccRecipientsSent)} (Non visible par les destinataires To/Cc)", cancellationToken);
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