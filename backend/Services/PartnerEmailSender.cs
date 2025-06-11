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
// FuzzySharp n'est définitivement plus nécessaire.
using ExcelFlow.Hubs;

namespace ExcelFlow.Services;

public class PartnerEmailSender
{
    private readonly SendEmail _sendEmailService;
    private readonly IHubContext<PartnerFileHub> _hubContext;

    public PartnerEmailSender(SendEmail sendEmailService, IHubContext<PartnerFileHub> hubContext)
    {
        _sendEmailService = sendEmailService;
        _hubContext = hubContext;
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
        // Ces chaînes seront les noms/sigles en minuscules, sans autre nettoyage.
        public string SearchableNameFull { get; set; } = string.Empty; // ex: "cecec (cec)" pour "CECEC (CEC)"
        public string? SearchableNameSigle { get; set; } = null;       // ex: "cec" pour "CECEC (CEC)"
    }

    public class SentEmailSummary
    {
        public string FileName { get; set; } = string.Empty;
        public string PartnerName { get; set; } = string.Empty;
        public List<string> RecipientEmails { get; set; } = new List<string>();
    }

    /// <summary>
    /// Normalise une chaîne en la convertissant en minuscules et en supprimant les accents.
    /// Les caractères spéciaux et les espaces sont conservés.
    /// </summary>
    /// <param name="input">La chaîne à normaliser.</param>
    /// <returns>La chaîne normalisée.</returns>
    private string NormalizeForComparison(string input)
    {
        input = input.ToLowerInvariant(); // Convertir en minuscules
        input = Regex.Replace(input.Normalize(System.Text.NormalizationForm.FormD), @"\p{M}", ""); // Supprimer les accents
        return input;
    }

    /// <summary>
    /// Lit les informations des partenaires (Nom, Emails) à partir d'un fichier Excel.
    /// </summary>
    /// <param name="partnerEmailFilePath">Chemin complet vers le fichier Excel.</param>
    /// <param name="cancellationToken">Token d'annulation.</param>
    /// <returns>Une liste de PartnerInfo.</returns>
    /// <exception cref="FileNotFoundException">Lancée si le fichier Excel n'est pas trouvé.</exception>
    /// <exception cref="InvalidOperationException">Lancée si le format du fichier Excel est incorrect, vide, ou les en-têtes requis sont manquants.</exception>
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
                                // Normaliser le nom complet du partenaire et potentiellement son sigle
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

    /// <summary>
    /// Envoie des emails à plusieurs partenaires, en analysant d'abord les fichiers générés pour identifier les destinataires.
    /// Chaque fichier trouvé est envoyé au(x) partenaire(s) dont le nom de fichier contient le nom du partenaire comme une sous-chaîne exacte de "mot entier",
    /// après normalisation (minuscules, sans accents).
    /// Un fichier ne peut être envoyé qu'à un seul partenaire.
    /// </summary>
    /// <param name="partnerEmailFilePath">Chemin complet vers le fichier Excel contenant les partenaires.</param>
    /// <param name="generatedFilesFolderPath">Chemin complet vers le dossier contenant les fichiers générés (pièces jointes).</param>
    /// <param name="subject">Sujet de l'email.</param>
    /// <param name="body">Corps de l'email (HTML).</param>
    /// <param name="fromDisplayName">Nom d'affichage de l'expéditeur. (Optionnel)</param>
    /// <param name="cancellationToken">Token to observe for cancellation requests.</param>
    /// <returns>Une tâche représentant l'opération d'envoi.</returns>
    public async Task SendEmailsToPartnersWithAttachments(
        string partnerEmailFilePath,
        string generatedFilesFolderPath,
        string subject,
        string body,
        string? fromDisplayName = null,
        CancellationToken cancellationToken = default)
    {
        await LogAndSend($"Démarrage du processus d'envoi d'emails basé sur les fichiers générés...", cancellationToken);

        // --- PARTIE LECTURE DES PARTENAIRES ---
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
                    await LogAndSend($"  - Partenaire: '{p.PartnerName}' | Emails: {string.Join(", ", p.Emails)} | Noms de recherche normalisés: {searchableNames}", cancellationToken);
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

        // --- PARTIE ANALYSE DES FICHIERS GÉNÉRÉS ---
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

        // --- Suivi des fichiers déjà affectés et récapitulatif ---
        var processedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sentEmailSummaries = new List<SentEmailSummary>(); // Pour le récapitulatif final

        // --- PARTIE TRAITEMENT ET ENVOI DES EMAILS ---
        int filesProcessedCount = 0;
        foreach (var filePath in allGeneratedFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (processedFiles.Contains(filePath))
            {
                await LogAndSend($"  Fichier '{Path.GetFileName(filePath)}' déjà traité pour un autre partenaire. Ignoré.", cancellationToken);
                await LogAndSend("---", cancellationToken);
                continue;
            }

            filesProcessedCount++;
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            await LogAndSend($"Traitement du fichier {filesProcessedCount}/{allGeneratedFiles.Count}: '{Path.GetFileName(filePath)}'", cancellationToken);

            // Normaliser le nom du fichier pour la comparaison (minuscules, sans accents)
            string normalizedFileName = NormalizeForComparison(fileNameWithoutExtension);

            PartnerInfo? foundPartner = null;

            // Parcourir les partenaires pour trouver une correspondance exacte avec frontière de mot
            foreach (var partner in partners)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // DÉFINITION DE LA CORRESPONDANCE
                // Utiliser Regex avec \b (frontière de mot) pour une correspondance exacte de mot/séquence.
                // Regex.Escape est crucial au cas où le nom du partenaire contiendrait des caractères spéciaux regex.
                // RegexOptions.CultureInvariant pour des correspondances indépendantes des paramètres culturels.
                // RegexOptions.IgnoreCase est géré par le NormalizeForComparison en amont.

                bool matched = false;

                // Tenter de faire correspondre le nom complet
                if (!string.IsNullOrWhiteSpace(partner.SearchableNameFull))
                {
                    // La regex \b{pattern}\b cherche le pattern comme un mot entier.
                    // Une "frontière de mot" (\b) est la position entre un caractère alphanumérique ou underscore (\w)
                    // et un caractère non-alphanumérique/non-underscore (\W), ou le début/fin de la chaîne.
                    var regexFull = new Regex($@"\b{Regex.Escape(partner.SearchableNameFull)}\b", RegexOptions.CultureInvariant);
                    if (regexFull.IsMatch(normalizedFileName))
                    {
                        matched = true;
                    }
                }

                // Si pas de match avec le nom complet et qu'un sigle existe, tenter de matcher le sigle
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
                    break; // Un fichier ne doit être envoyé qu'à un seul partenaire. Le premier match "mot entier" gagne.
                }
            }

            if (foundPartner != null && foundPartner.Emails.Any())
            {
                await LogAndSend($"  Fichier '{Path.GetFileName(filePath)}' contient le nom/sigle exact et ordonné du partenaire '{foundPartner.PartnerName}'.", cancellationToken);
                await LogAndSend($"  Adresses email cibles pour '{foundPartner.PartnerName}': {string.Join(", ", foundPartner.Emails)}", cancellationToken);
                await LogAndSend($"  Envoi de l'email à {foundPartner.PartnerName} avec le fichier '{Path.GetFileName(filePath)}' en pièce jointe.", cancellationToken);

                bool sent = await _sendEmailService.SendEmailAsync(
                    subject: subject,
                    body: body,
                    toRecipients: foundPartner.Emails,
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
                        RecipientEmails = foundPartner.Emails.ToList()
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
                    await LogAndSendError($"  Partenaire '{foundPartner.PartnerName}' trouvé pour le fichier '{Path.GetFileName(filePath)}' mais sans adresse email valide. Email non envoyé.", cancellationToken);
                }
                else
                {
                    await LogAndSend($"  Aucun partenaire dont le nom/sigle exact et ordonné est inclus comme un mot entier dans le fichier '{Path.GetFileName(filePath)}'. Le fichier sera ignoré.", cancellationToken);
                }
            }
            await LogAndSend("---", cancellationToken);
        }
        await LogAndSend("Processus d'envoi d'emails basé sur les fichiers générés terminé.", cancellationToken);

        // --- Récapitulatif final ---
        await LogAndSend("\n--- RÉCAPITULATIF DES EMAILS ENVOYÉS ---", cancellationToken);
        if (sentEmailSummaries.Any())
        {
            await LogAndSend($"Total de fichiers envoyés : {sentEmailSummaries.Count}", cancellationToken);
            foreach (var summary in sentEmailSummaries)
            {
                await LogAndSend($"  - Fichier: '{summary.FileName}' envoyé à Partenaire: '{summary.PartnerName}' ({string.Join(", ", summary.RecipientEmails)})", cancellationToken);
            }
        }
        else
        {
            await LogAndSend("Aucun email n'a été envoyé durant ce processus.", cancellationToken);
        }
        await LogAndSend("---------------------------------------", cancellationToken);
    }
}