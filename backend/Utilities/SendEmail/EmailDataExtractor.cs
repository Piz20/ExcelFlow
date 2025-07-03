// Fichier : Utilities/EmailDataExtractor.cs
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
using ExcelFlow.Models; 

namespace ExcelFlow.Utilities
{
    public static class StringExtensions
    {
        public static string NormalizeSpaces(this string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // Remplace espaces insécables (U+00A0), tabulations, multiples espaces, etc. par un espace simple
            var cleaned = Regex.Replace(input, @"[\u00A0\s]+", " ");

            // Trim classique
            return cleaned.Trim();
        }
    }

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

        public async Task<EmailData?> ExtractEmailDataFromAttachment(string filePath, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!File.Exists(filePath))
            {
                await _logAndSendError($"Le fichier pièce jointe est introuvable : {filePath}", cancellationToken);
                return null;
            }

            var data = new EmailData(); 
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string fileNameCleaned = fileName.NormalizeSpaces();

            await _logAndSend($"Analyse du fichier Excel pour l'extraction de données : '{Path.GetFileName(filePath)}'", cancellationToken);

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

                    // --- 1. Extraire la date/intervalle du nom du fichier ---
                    string spaceOrSeparator = @"[\s\u00A0./]";
                    string datePattern = $@"\d{{2}}{spaceOrSeparator}\d{{2}}{spaceOrSeparator}(?:\d{{4}}|\d{{2}})";
                    var dateMatch = Regex.Match(fileNameCleaned, $@"du\s+({datePattern})(?:\s+au\s+({datePattern}))?", RegexOptions.IgnoreCase);

                    if (dateMatch.Success)
                    {
                        string startDate = dateMatch.Groups[1].Value.NormalizeSpaces();
                        string endDate = dateMatch.Groups[2].Success ? dateMatch.Groups[2].Value.NormalizeSpaces() : string.Empty;

                        if (!string.IsNullOrEmpty(endDate))
                        {
                            data.DateString = $"{startDate} au {endDate}";
                        }
                        else
                        {
                            data.DateString = startDate;
                        }
                        await _logAndSend($"Date/Intervalle extrait du nom de fichier : '{data.DateString}'", cancellationToken);
                    }
                    else
                    {
                        await _logAndSendError($"Impossible d'extraire la date ou l'intervalle du nom de fichier : '{fileName}'. Les formats 'du JJ.MM.AAAA', 'du JJ/MM/AAAA', 'du JJ MM AAAA' (ou AA) ou leurs intervalles n'ont pas été trouvés.", cancellationToken);
                        data.DateString = "[DATE_NON_TROUVÉE]";
                    }

                    // --- 2. Extraire le nom du partenaire : première cellule non vide sous l'en-tête "PARTENAIRES" ---
                    const string partnerNameHeader = "PARTENAIRES"; 
                    const string headerToFindBalance = "Solde fin de journée avant Cash out Auto";

                    int partnerNameCol = -1;
                    int balanceCol = -1;
                    int headerRow = -1;

                    // Scan up to the first 10 rows to find headers
                    var lastRowUsedInHeaders = worksheet.LastRowUsed();
                    int maxHeaderRowsToScan = Math.Min(10, lastRowUsedInHeaders != null ? lastRowUsedInHeaders.RowNumber() : worksheet.LastRow().RowNumber());

                    for (int rowNum = 1; rowNum <= maxHeaderRowsToScan; rowNum++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var row = worksheet.Row(rowNum);
                        foreach (var cell in row.CellsUsed())
                        {
                            string headerText = cell.Value.ToString()?.NormalizeSpaces() ?? string.Empty;

                            if (!string.IsNullOrEmpty(headerText))
                            {
                                if (headerText.Equals(partnerNameHeader, StringComparison.OrdinalIgnoreCase))
                                {
                                    partnerNameCol = cell.Address.ColumnNumber;
                                    headerRow = rowNum; 
                                }
                                else if (headerText.Equals(headerToFindBalance, StringComparison.OrdinalIgnoreCase))
                                {
                                    balanceCol = cell.Address.ColumnNumber;
                                    if (headerRow == -1) headerRow = rowNum; 
                                }
                            }

                            if (partnerNameCol != -1 && balanceCol != -1 && headerRow != -1)
                                break;
                        }
                        if (partnerNameCol != -1 && balanceCol != -1 && headerRow != -1)
                            break;
                    }
                    
                    if (partnerNameCol == -1)
                    {
                        await _logAndSendError($"L'en-tête obligatoire '{partnerNameHeader}' n'a pas été trouvé dans le fichier Excel '{Path.GetFileName(filePath)}'.", cancellationToken);
                        data.PartnerNameInFile = "[NOM_PARTENAIRE_NON_TROUVÉ]";
                    }
                    else
                    {
                        IXLCell? partnerNameCell = null;
                        // On commence à chercher la première cellule non vide APRÈS la ligne d'en-tête
                        var lastRowUsedInWorksheet = worksheet.LastRowUsed();
                        int startScanRow = headerRow + 1;

                        if (lastRowUsedInWorksheet != null)
                        {
                            for (int rowNum = startScanRow; rowNum <= lastRowUsedInWorksheet.RowNumber(); rowNum++)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                var currentCell = worksheet.Cell(rowNum, partnerNameCol);
                                if (!currentCell.IsEmpty())
                                {
                                    partnerNameCell = currentCell;
                                    break; // On a trouvé la première cellule non vide
                                }
                            }
                        }

                        if (partnerNameCell != null)
                        {
                            data.PartnerNameInFile = partnerNameCell.Value.ToString().NormalizeSpaces();
                            await _logAndSend($"Nom du partenaire extrait du fichier Excel ('{partnerNameCell.Address.ToString()}') : '{data.PartnerNameInFile}'", cancellationToken);
                        }
                        else
                        {
                            await _logAndSendError($"Aucune cellule non vide trouvée sous l'en-tête '{partnerNameHeader}' dans le fichier '{Path.GetFileName(filePath)}'. Le nom du partenaire n'a pas pu être extrait.", cancellationToken);
                            data.PartnerNameInFile = "[NOM_PARTENAIRE_NON_TROUVÉ]";
                        }
                    }

                    // --- 3. Extraire le solde final du fichier Excel (logique inchangée) ---
                    if (balanceCol == -1)
                    {
                        await _logAndSendError($"L'en-tête '{headerToFindBalance}' n'a pas été trouvé dans le fichier Excel '{Path.GetFileName(filePath)}'.", cancellationToken);
                        data.FinalBalance = "[SOLDE_NON_TROUVÉ]";
                        return data; 
                    }

                    await _logAndSend($"En-tête '{headerToFindBalance}' trouvé dans la colonne {balanceCol} (ligne {headerRow}). Recherche du solde final...", cancellationToken);

                    decimal? lastFoundBalance = null;
                    IXLCell? lastFoundCell = null;

                    var dataLastRowUsed = worksheet.LastRowUsed();
                    int startDataRowForBalance = headerRow + 1; // Start scanning for balance after the header row

                    if (dataLastRowUsed != null)
                    {
                        for (int rowNum = startDataRowForBalance; rowNum <= dataLastRowUsed.RowNumber(); rowNum++)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            var cell = worksheet.Cell(rowNum, balanceCol);

                            if (!cell.IsEmpty())
                            {
                                if (cell.TryGetValue(out decimal cellValue))
                                {
                                    lastFoundBalance = cellValue;
                                    lastFoundCell = cell;
                                    await _logAndSend($"    Valeur numérique trouvée à {cell.Address.ToString()}: {cellValue}", cancellationToken);
                                }
                                else if (decimal.TryParse(cell.Value.ToString()?.NormalizeSpaces().Replace('.', ','), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("fr-FR"), out cellValue))
                                {
                                    lastFoundBalance = cellValue;
                                    lastFoundCell = cell;
                                    await _logAndSend($"    Valeur chaîne numérique trouvée (parsed) à {cell.Address.ToString()}: {cellValue}", cancellationToken);
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
                        await _logAndSendError($"Aucune valeur numérique trouvée sous l'en-tête '{headerToFindBalance}' dans le fichier '{Path.GetFileName(filePath)}'.", cancellationToken);
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
                await _logAndSendError($"Erreur inattendue lors de l'extraction des données du fichier Excel '{Path.GetFileName(filePath)}' : {ex.Message}", cancellationToken);
                return null;
            }

            return data;
        }
    }
}
