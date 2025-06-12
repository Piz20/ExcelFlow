// Fichier : Services/PartnerExcelReader.cs
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ExcelFlow.Models; // Assurez-vous d'ajouter cette ligne

namespace ExcelFlow.Services;

public class PartnerExcelReader
{
    public string NormalizeForComparison(string input)
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
                    string headerText = cell.Value.ToString()?.Trim() ?? string.Empty;
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
}