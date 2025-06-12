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
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.SignalR;
using System.Threading;
using ExcelFlow.Hubs;
using ExcelFlow.Utilities;

namespace ExcelFlow.Services;

public class PartnerEmailSender
{
    private readonly SendEmail _sendEmailService;
    private readonly IHubContext<PartnerFileHub> _hubContext;
    private readonly EmailDataExtractor _emailDataExtractor;

    // --- TEMPLATES HARDCODED HERE ---
    private const string SUBJECT_TEMPLATE = "Objet : Détermination du solde final de [NOM_PARTENAIRE] - [DATE_FICHIER]";
    // --- NOUVEAU BODY_TEMPLATE avec HTML ---
    private const string BODY_TEMPLATE = @"
<p>Bonsoir cher partenaire,</p>
<p>Merci de trouver en pièce jointe ci-dessous l'analyse de votre compte support pour les journées du <strong>[DATE_OU_INTERVALLE_JOURS_ANALYSE]</strong>, ayant permis d'aboutir au solde final (avant cash out auto) de <span style='color: #f51b1b;'><strong>[SOLDE_FINAL] [CURRENCY]</strong></span>.</p>
<p>Vous trouverez également les éléments ayant servi de base au calcul de votre solde final, pour les journées du <strong>[DATE_OU_INTERVALLE_JOURS_ANALYSE]</strong>.</p>
<p>Nous restons disponibles pour toute information complémentaire.</p>
<p><span style='color: #f51b1b;'><strong>NB :</strong> Prière d'effectuer vos contrôles caisse et compte support à J+1 (dans les 24H) afin de nous remonter toute anomalie constatée pour régularisation, soit à notre niveau, soit au niveau du MTO.</span></p><p>Merci d'accuser réception.</p>
<p>Cordialement,</p>";
    // ---------------------------------

    public PartnerEmailSender(SendEmail sendEmailService, IHubContext<PartnerFileHub> hubContext, EmailDataExtractor emailDataExtractor)
    {
        _sendEmailService = sendEmailService;
        _hubContext = hubContext;
        _emailDataExtractor = emailDataExtractor;
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

    public class PartnerInfo
    {
        public string PartnerName { get; set; } = string.Empty;
        public List<string> Emails { get; set; } = new List<string>();
        public string SearchableNameFull { get; set; } = string.Empty;
        public string? SearchableNameSigle { get; set; } = null;
    }

    public class SentEmailSummary
    {
        public string FileName { get; set; } = string.Empty;
        public string PartnerName { get; set; } = string.Empty;
        public List<string> RecipientEmails { get; set; } = new List<string>();
        public List<string> CcRecipientsSent { get; set; } = new List<string>();
        public List<string> BccRecipientsSent { get; set; } = new List<string>();
    }

    private string NormalizeForComparison(string input)
    {
        input = input.ToLowerInvariant();
        input = Regex.Replace(input.Normalize(System.Text.NormalizationForm.FormD), @"\p{M}", "");
        return input;
    }

    public List<PartnerInfo> ReadPartnersFromExcel(string partnerEmailFilePath, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!File.Exists(partnerEmailFilePath))
        {
            throw new FileNotFoundException($"Le fichier Excel est introuvable : {partnerEmailFilePath}");
        }

        var partners = new List<PartnerInfo>();

        using (var workbook = new XLWorkbook(partnerEmailFilePath))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var worksheet = workbook.Worksheets.FirstOrDefault();
            if (worksheet == null || worksheet.RowsUsed().Count() < 2)
            {
                throw new InvalidOperationException("Le fichier Excel ne contient aucune feuille de calcul ou est vide.");
            }

            const string partnerNameHeader = "NOM DU PARTENAIRE";
            const string emailHeader = "ADRESSES";

            int partnerNameCol = -1;
            int emailCol = -1;
            int headerRow = -1;

            var lastRowUsed = worksheet.LastRowUsed();
            int lastRowNumber = lastRowUsed != null ? lastRowUsed.RowNumber() : worksheet.LastRow().RowNumber();
            for (int rowNum = 1; rowNum <= Math.Min(10, lastRowNumber); rowNum++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var row = worksheet.Row(rowNum);
                foreach (var cell in row.CellsUsed())
                {
                    string headerText = cell.Value.ToString() != null ? cell.Value.ToString().Trim() : string.Empty;
                    if (!string.IsNullOrEmpty(headerText) && headerText.Equals(partnerNameHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        partnerNameCol = cell.Address.ColumnNumber;
                        headerRow = rowNum;
                    }
                    else if (!string.IsNullOrEmpty(headerText) && headerText.Equals(emailHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        emailCol = cell.Address.ColumnNumber;
                        headerRow = rowNum;
                    }

                    if (partnerNameCol != -1 && emailCol != -1 && headerRow != -1)
                        break;
                }
                if (partnerNameCol != -1 && emailCol != -1 && headerRow != -1)
                    break;
            }

            if (partnerNameCol == -1 || emailCol == -1)
            {
                throw new InvalidOperationException($"Les en-têtes requis '{partnerNameHeader}' ou '{emailHeader}' n'ont pas été trouvés dans le fichier Excel.");
            }

            var emailRegex = new Regex(@"([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?)", RegexOptions.IgnoreCase);

            var dataLastRowUsed = worksheet.LastRowUsed();
            if (dataLastRowUsed != null)
            {
                for (int row = headerRow + 1; row <= dataLastRowUsed.RowNumber(); row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var partnerNameCell = worksheet.Cell(row, partnerNameCol).Value;
                    var emailCell = worksheet.Cell(row, emailCol).Value;

                    if (!partnerNameCell.IsBlank)
                    {
                        string partnerName = partnerNameCell.ToString()?.Trim() ?? string.Empty;
                        List<string> extractedEmails = new List<string>();

                        if (!emailCell.IsBlank)
                        {
                            string emailsString = emailCell.ToString()?.Trim() ?? string.Empty;
                            foreach (Match match in emailRegex.Matches(emailsString))
                            {
                                if (!string.IsNullOrWhiteSpace(match.Groups[1].Value))
                                {
                                    extractedEmails.Add(match.Groups[1].Value.Trim());
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(partnerName) && extractedEmails.Any())
                        {
                            var existingPartner = partners.FirstOrDefault(p => p.PartnerName.Equals(partnerName, StringComparison.OrdinalIgnoreCase));
                            if (existingPartner != null)
                            {
                                foreach (var email in extractedEmails)
                                {
                                    if (!existingPartner.Emails.Contains(email, StringComparer.OrdinalIgnoreCase))
                                    {
                                        existingPartner.Emails.Add(email);
                                    }
                                }
                            }
                            else
                            {
                                string searchableNameFull = NormalizeForComparison(partnerName);
                                string? searchableNameSigle = null;

                                Match sigleMatch = Regex.Match(partnerName, @"\(([^)]+)\)");
                                if (sigleMatch.Success && !string.IsNullOrWhiteSpace(sigleMatch.Groups[1].Value))
                                {
                                    searchableNameSigle = NormalizeForComparison(sigleMatch.Groups[1].Value.Trim());
                                }

                                partners.Add(new PartnerInfo
                                {
                                    PartnerName = partnerName,
                                    Emails = extractedEmails.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                                    SearchableNameFull = searchableNameFull,
                                    SearchableNameSigle = searchableNameSigle
                                });
                            }
                        }
                    }
                }
            }
        }
        return partners;
    }

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
            partners = await Task.Run(() => ReadPartnersFromExcel(partnerEmailFilePath, cancellationToken), cancellationToken);
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

            string normalizedFileName = NormalizeForComparison(fileNameWithoutExtension);

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

                var emailData = await _emailDataExtractor.ExtractEmailDataFromAttachment(filePath, cancellationToken);
                if (emailData == null)
                {
                    await LogAndSendError($"Impossible d'extraire les données dynamiques pour l'email du fichier '{Path.GetFileName(filePath)}'. Email non envoyé.", cancellationToken);
                    await LogAndSend("---", cancellationToken);
                    continue;
                }

                // --- Dynamic subject and body formatting using hardcoded templates ---
                string finalSubject = SUBJECT_TEMPLATE
                    .Replace("[NOM_PARTENAIRE]", emailData.PartnerNameInFile)
                    .Replace("[DATE_FICHIER]", emailData.DateString);

                string finalBody = BODY_TEMPLATE
                    .Replace("[DATE_OU_INTERVALLE_JOURS_ANALYSE]", emailData.DateString)
                    .Replace("[SOLDE_FINAL]", emailData.FinalBalance)
                    .Replace("[CURRENCY]", emailData.Currency);


                await LogAndSend($"   Envoi de l'email à {foundPartner.PartnerName} avec le fichier '{Path.GetFileName(filePath)}' en pièce jointe.", cancellationToken);

                // --- APPEL CORRECT DE SendEmailAsync SANS 'isHtml' ---
                bool sent = await _sendEmailService.SendEmailAsync(
                    subject: finalSubject,
                    body: finalBody,
                    toRecipients: foundPartner.Emails,
                    ccRecipients: ccRecipients,
                    bccRecipients: bccRecipients,
                    fromDisplayName: fromDisplayName,
                    attachmentFilePaths: new List<string> { filePath },
                    cancellationToken: cancellationToken // Maintenant, le CancellationToken est le dernier paramètre
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