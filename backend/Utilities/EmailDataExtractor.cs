using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.SignalR;
using ExcelFlow.Hubs;

namespace ExcelFlow.Utilities;

public class EmailDataExtractor
{
    private readonly IHubContext<PartnerFileHub>? _hubContext;
    private readonly Func<string, CancellationToken, Task> _logAndSend;
    private readonly Func<string, CancellationToken, Task> _logAndSendError;

    public EmailDataExtractor(IHubContext<PartnerFileHub>? hubContext = null)
    {
        _hubContext = hubContext;
        _logAndSend = async (message, token) =>
        {
            string formattedMessage = $"[{DateTime.Now:HH:mm:ss}] {message}";
            Console.WriteLine(formattedMessage);
            if (_hubContext != null)
            {
                await _hubContext.Clients.All.SendAsync("ReceiveMessage", formattedMessage, token);
            }
        };
        _logAndSendError = async (message, token) =>
        {
            string formattedMessage = $"[{DateTime.Now:HH:mm:ss}] ERREUR: {message}";
            Console.Error.WriteLine(formattedMessage);
            if (_hubContext != null)
            {
                await _hubContext.Clients.All.SendAsync("ReceiveErrorMessage", formattedMessage, token);
            }
        };
    }

    public EmailDataExtractor(Func<string, CancellationToken, Task> logAndSend, Func<string, CancellationToken, Task> logAndSendError)
    {
        _logAndSend = logAndSend ?? ((msg, token) => { Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {msg}"); return Task.CompletedTask; });
        _logAndSendError = logAndSendError ?? ((msg, token) => { Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] ERREUR: {msg}"); return Task.CompletedTask; });
    }

    public class ExtractedEmailData
    {
        public string DateString { get; set; } = string.Empty;
        public string FinalBalance { get; set; } = string.Empty;
        public string Currency { get; set; } = "XAF";
        public string PartnerNameInFile { get; set; } = string.Empty;
    }

    public async Task<ExtractedEmailData?> ExtractEmailDataFromAttachment(string filePath, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!File.Exists(filePath))
        {
            await _logAndSendError($"Le fichier pièce jointe est introuvable : {filePath}", cancellationToken);
            return null;
        }

        var data = new ExtractedEmailData();
        string fileName = Path.GetFileNameWithoutExtension(filePath);
        await _logAndSend($"Analyse du nom de fichier pour extraction : '{fileName}'", cancellationToken);

        // --- 1. Extraire la date/intervalle du nom du fichier ---
        // Nouveau pattern pour gérer tous les formats de date (JJ MM AA/AAAA, JJ.MM.AA/AAAA, JJ/MM/AA/AAAA)
        // en utilisant '[\s./]' pour matcher espace, point ou slash comme séparateur.
        string datePattern = @"\d{2}[\s./]\d{2}[\s./](?:\d{4}|\d{2})"; // JJ.MM.AAAA ou JJ.MM.AA
        var dateMatch = Regex.Match(fileName, $@"du\s+({datePattern})(?:\s+au\s+({datePattern}))?", RegexOptions.IgnoreCase);

        if (dateMatch.Success)
        {
            string startDate = dateMatch.Groups[1].Value.Trim();
            string endDate = dateMatch.Groups[2].Success ? dateMatch.Groups[2].Value.Trim() : string.Empty;

            if (!string.IsNullOrEmpty(endDate))
            {
                data.DateString = $"{startDate} au {endDate}";
            }
            else
            {
                data.DateString = startDate;
            }
            await _logAndSend($"Date/Intervalle extrait : '{data.DateString}'", cancellationToken);
        }
        else
        {
            await _logAndSendError($"Impossible d'extraire la date ou l'intervalle du nom de fichier : '{fileName}'. Les formats 'du JJ.MM.AAAA', 'du JJ/MM/AAAA', 'du JJ MM AAAA' (ou AA) ou leurs intervalles n'ont pas été trouvés.", cancellationToken);
            data.DateString = "[DATE_NON_TROUVÉE]";
        }

        // --- 2. Extraire le nom du partenaire du nom du fichier ---
        // Le pattern précédent est maintenu.
        var partnerNameMatch = Regex.Match(fileName, @"COMPTE SUPPORT\s+([^d]+)\s+du", RegexOptions.IgnoreCase);

        if (partnerNameMatch.Success && partnerNameMatch.Groups.Count > 1)
        {
            data.PartnerNameInFile = partnerNameMatch.Groups[1].Value.Trim();
            await _logAndSend($"Nom du partenaire extrait : '{data.PartnerNameInFile}'", cancellationToken);
        }
        else
        {
            await _logAndSendError($"Impossible d'extraire le nom du partenaire du nom de fichier : '{fileName}'. Le format 'COMPTE SUPPORT [NOM_PARTENAIRE] du ...' n'a pas été trouvé.", cancellationToken);
            data.PartnerNameInFile = "[NOM_PARTENAIRE_NON_TROUVÉ]";
        }

        // --- 3. Extraire le solde final du fichier Excel (logique inchangée, elle était déjà robuste) ---
        try
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var worksheet = workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    await _logAndSendError($"Le fichier Excel '{Path.GetFileName(filePath)}' ne contient aucune feuille de calcul.", cancellationToken);
                    return null;
                }

                const string headerToFind = "Solde fin de journée avant Cash out Auto";
                int headerRow = -1;
                int targetColumn = -1;

                var lastRowUsedInHeaders = worksheet.LastRowUsed();
                int maxHeaderRowsToScan = Math.Min(10, lastRowUsedInHeaders != null ? lastRowUsedInHeaders.RowNumber() : worksheet.LastRow().RowNumber());

                for (int rowNum = 1; rowNum <= maxHeaderRowsToScan; rowNum++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var row = worksheet.Row(rowNum);
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.Value.ToString().Trim().Equals(headerToFind, StringComparison.OrdinalIgnoreCase))
                        {
                            targetColumn = cell.Address.ColumnNumber;
                            headerRow = rowNum;
                            break;
                        }
                    }
                    if (targetColumn != -1) break;
                }

                if (targetColumn == -1)
                {
                    await _logAndSendError($"L'en-tête '{headerToFind}' n'a pas été trouvé dans le fichier Excel '{Path.GetFileName(filePath)}'.", cancellationToken);
                    data.FinalBalance = "[SOLDE_NON_TROUVÉ]";
                    return data;
                }

                await _logAndSend($"En-tête '{headerToFind}' trouvé dans la colonne {targetColumn} (ligne {headerRow}). Recherche du solde final...", cancellationToken);

                decimal? lastFoundBalance = null;
                IXLCell? lastFoundCell = null;

                var dataLastRowUsed = worksheet.LastRowUsed();
                int startDataRow = headerRow + 1;

                if (dataLastRowUsed != null)
                {
                    for (int rowNum = startDataRow; rowNum <= dataLastRowUsed.RowNumber(); rowNum++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var cell = worksheet.Cell(rowNum, targetColumn);

                        if (!cell.IsEmpty())
                        {
                            if (cell.TryGetValue(out decimal cellValue))
                            {
                                lastFoundBalance = cellValue;
                                lastFoundCell = cell;
                                await _logAndSend($"   Valeur numérique trouvée à {cell.Address.ToString()}: {cellValue}", cancellationToken);
                            }
                            else if (decimal.TryParse(cell.Value.ToString()?.Trim().Replace('.', ','), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("fr-FR"), out cellValue))
                            {
                                lastFoundBalance = cellValue;
                                lastFoundCell = cell;
                                await _logAndSend($"   Valeur chaîne numérique trouvée (parsed) à {cell.Address.ToString()}: {cellValue}", cancellationToken);
                            }
                        }
                    }
                }

                if (lastFoundBalance.HasValue)
                {
                    data.FinalBalance = lastFoundBalance.Value.ToString("#,##0.00", System.Globalization.CultureInfo.GetCultureInfo("fr-FR"));
                    await _logAndSend($"Solde final extrait : {data.FinalBalance} de la cellule {lastFoundCell?.Address.ToString() ?? "N/A"}", cancellationToken);
                }
                else
                {
                    await _logAndSendError($"Aucune valeur numérique trouvée sous l'en-tête '{headerToFind}' dans le fichier '{Path.GetFileName(filePath)}'.", cancellationToken);
                    data.FinalBalance = "[SOLDE_NON_TROUVÉ]";
                }
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            await _logAndSendError($"Erreur lors de l'extraction du solde final du fichier Excel '{Path.GetFileName(filePath)}' : {ex.Message}", cancellationToken);
            return null;
        }

        return data;
    }
}