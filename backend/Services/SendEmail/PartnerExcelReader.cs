// Fichier : Services/PartnerExcelReader.cs
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ExcelFlow.Models; 

namespace ExcelFlow.Services;

public class PartnerExcelReader
{
    public string NormalizeForComparison(string input)
    {
        input = input.ToLowerInvariant();
        input = Regex.Replace(input.Normalize(System.Text.NormalizationForm.FormD), @"\p{M}", "");
        return input;
    }

    public List<string> ExtractSearchableKeywords(string partnerName)
    {
        var keywords = new List<string>();
        var parts = Regex.Split(partnerName, @"[\s\-_.,()]+").Where(p => !string.IsNullOrWhiteSpace(p)).ToArray();

        var stopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "societe", "société", "groupe", "group", "sa", "sarl", "sas", "inc", "llc", "ltd", "co", "company",
            "agence", "entreprise", "holding", "management", "international", "national", "global", "de", "du", "des", "la", "le", "les", "et", "ou", "a",
            "plc", "corp", "corporation", "ltd", "limited", "s.a", "s.p.a", "gmbh", "ag", "s.a.r.l", "s.a.s"
        };

        foreach (var part in parts)
        {
            var normalizedPart = NormalizeForComparison(part);
            if (!string.IsNullOrWhiteSpace(normalizedPart) && !stopWords.Contains(normalizedPart) && normalizedPart.Length > 1)
            {
                keywords.Add(normalizedPart);
            }
        }

        Match sigleMatch = Regex.Match(partnerName, @"\(([^)]+)\)");
        if (sigleMatch.Success && !string.IsNullOrWhiteSpace(sigleMatch.Groups[1].Value))
        {
            var normalizedSigle = NormalizeForComparison(sigleMatch.Groups[1].Value.Trim());
            if (!string.IsNullOrWhiteSpace(normalizedSigle) && !keywords.Contains(normalizedSigle))
            {
                keywords.Add(normalizedSigle);
            }
        }
        
        var normalizedFullName = NormalizeForComparison(partnerName);
        if (!string.IsNullOrWhiteSpace(normalizedFullName) && !keywords.Contains(normalizedFullName))
        {
            keywords.Add(normalizedFullName);
        }

        return keywords.Distinct().ToList();
    }

    /// <summary>
    /// Lit toutes les informations des partenaires à partir d'un fichier Excel.
    /// Chaque ligne de données représente un partenaire, avec son nom tiré de la colonne "PARTENAIRES".
    /// </summary>
    /// <param name="partnerEmailFilePath">Le chemin vers le fichier Excel contenant les partenaires.</param>
    /// <param name="cancellationToken">Token d'annulation.</param>
    /// <returns>Une liste d'objets PartnerInfo.</returns>
    /// <exception cref="FileNotFoundException">Si le fichier n'existe pas.</exception>
    /// <exception cref="InvalidOperationException">Si le fichier est mal formaté (en-têtes manquants).</exception>
    public List<PartnerInfo> ReadPartnersFromExcel(string partnerEmailFilePath, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!File.Exists(partnerEmailFilePath))
        {
            throw new FileNotFoundException($"Le fichier Excel est introuvable : {partnerEmailFilePath}");
        }

        List<PartnerInfo> partners = new List<PartnerInfo>();

        using (var workbook = new XLWorkbook(partnerEmailFilePath))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var worksheet = workbook.Worksheets.FirstOrDefault();
            if (worksheet == null || worksheet.RowsUsed().Count() < 2)
            {
                // Si le fichier est vide ou n'a qu'un en-tête, retourner une liste vide.
                return partners;
            }

            // Correction ici : Ajout de la première constante pour "NOM DU PARTENAIRE"
            const string partnerNameHeader1 = "NOM DU PARTENAIRE";
            const string partnerNameHeader2 = "PARTENAIRES"; // Deuxième en-tête possible
            const string emailHeader = "ADRESSES";

            int partnerNameCol = -1;
            int emailCol = -1;
            int headerRow = -1;

            // Recherche des en-têtes dans les 10 premières lignes
            var lastRowUsedInHeaderSearch = worksheet.LastRowUsed();
            int lastRowNumberForHeaderSearch = lastRowUsedInHeaderSearch != null ? lastRowUsedInHeaderSearch.RowNumber() : worksheet.LastRow().RowNumber();
            for (int rowNum = 1; rowNum <= Math.Min(10, lastRowNumberForHeaderSearch); rowNum++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var row = worksheet.Row(rowNum);
                foreach (var cell in row.CellsUsed())
                {
                    string headerText = cell.Value.ToString()?.Trim() ?? string.Empty;
                    // Vérifie si l'en-tête correspond à l'un des deux noms possibles (ignorant la casse)
                    if (!string.IsNullOrEmpty(headerText) && 
                        (headerText.Equals(partnerNameHeader1, StringComparison.OrdinalIgnoreCase) ||
                         headerText.Equals(partnerNameHeader2, StringComparison.OrdinalIgnoreCase)))
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
                // Mise à jour du message d'erreur pour inclure les deux options d'en-tête
                throw new InvalidOperationException($"Les en-têtes requis ('{partnerNameHeader1}' ou '{partnerNameHeader2}') et '{emailHeader}' n'ont pas été trouvés dans le fichier Excel. Veuillez vérifier la structure du fichier.");
            }

            // Parcourir toutes les lignes de données après l'en-tête
            // Utilise RowsUsed().Skip() pour commencer après la ligne d'en-tête
            var dataRows = worksheet.RowsUsed().Skip(headerRow).ToList(); 
            if (!dataRows.Any()) 
            {
                return partners; // Aucune ligne de données trouvée
            }

            foreach (var row in dataRows)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // Ne pas prendre la première ligne qui est vide mais plutot la première ligne qui contienne des données
                if (row.IsEmpty()) continue;

                var partnerNameCell = row.Cell(partnerNameCol).Value;
                var emailCell = row.Cell(emailCol).Value;

                // Si le nom du partenaire ou l'email est vide, ignorer cette ligne
                if (partnerNameCell.IsBlank || emailCell.IsBlank)
                {
                    Console.WriteLine($"[AVERTISSEMENT] Ligne {row.RowNumber()} ignorée : Nom du partenaire ou adresse email manquant.");
                    continue; 
                }

                string partnerName = partnerNameCell.ToString()?.Trim() ?? string.Empty;
                List<string> extractedEmails = new List<string>();

                var emailRegex = new Regex(@"([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?)", RegexOptions.IgnoreCase);
                string emailsString = emailCell.ToString()?.Trim() ?? string.Empty;
                foreach (Match match in emailRegex.Matches(emailsString))
                {
                    if (!string.IsNullOrWhiteSpace(match.Groups[1].Value))
                    {
                        extractedEmails.Add(match.Groups[1].Value.Trim());
                    }
                }

                if (!extractedEmails.Any())
                {
                    Console.WriteLine($"[AVERTISSEMENT] Ligne {row.RowNumber()} ignorée : Aucune adresse email valide trouvée pour le partenaire '{partnerName}'.");
                    continue; 
                }

                string searchableNameFull = NormalizeForComparison(partnerName);
                string? searchableNameSigle = null;

                Match sigleMatch = Regex.Match(partnerName, @"\(([^)]+)\)");
                if (sigleMatch.Success && !string.IsNullOrWhiteSpace(sigleMatch.Groups[1].Value))
                {
                    searchableNameSigle = NormalizeForComparison(sigleMatch.Groups[1].Value.Trim());
                }

                List<string> searchableKeywords = ExtractSearchableKeywords(partnerName);

                partners.Add(new PartnerInfo
                {
                    PartnerName = partnerName,
                    Emails = extractedEmails.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                    SearchableNameFull = searchableNameFull,
                    SearchableNameSigle = searchableNameSigle,
                    SearchableKeywords = searchableKeywords
                });
            }
        }
        return partners;
    }
}