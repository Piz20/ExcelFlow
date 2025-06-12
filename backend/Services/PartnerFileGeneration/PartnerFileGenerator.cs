using ClosedXML.Excel;
using Microsoft.AspNetCore.SignalR;
using ExcelFlow.Hubs;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System;
using System.Linq; // Nécessaire pour .FirstOrDefault(), .Any()
using ExcelFlow.Models; // Pour utiliser ProgressUpdate
namespace ExcelFlow.Services
{
    public class PartnerFileGenerator
    {
        private readonly IHubContext<PartnerFileHub> _hubContext;

        public PartnerFileGenerator(IHubContext<PartnerFileHub> hubContext)
        {
            _hubContext = hubContext;
        }

        // MODIFICATION ICI : Nouvelle surcharge de LogAndSend
        // Cette version envoie le message à la console du serveur et au client.
        private async Task LogAndSend(string message, CancellationToken cancellationToken = default)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
            await _hubContext.Clients.All.SendAsync("ReceiveMessage", message, cancellationToken);
        }

        // NOUVELLE FONCTION : LogOnly
        // Cette fonction n'affiche le message qu'en console du serveur, PAS au client.
        private void LogOnly(string message)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }

        public async Task GeneratePartnerFilesAsync(
            IXLWorksheet worksheet,
            string templatePath,
            string outputDir,
            int startIndex = 0,
            int count = 3,
            CancellationToken cancellationToken = default)
        {
            await LogAndSend("--- Démarrage du processus de génération des fichiers partenaires ---", cancellationToken);

            // Vérifications initiales de la feuille
            var lastRowUsed = worksheet.LastRowUsed();
            if (lastRowUsed == null)
            {
                await LogAndSend("❌ Erreur: La feuille de calcul ne contient aucune ligne utilisée.", cancellationToken);
                throw new InvalidOperationException("La feuille de calcul ne contient aucune ligne utilisée.");
            }
            int lastRow = lastRowUsed.RowNumber();
            LogOnly($"Dernière ligne utilisée détectée : {lastRow}"); // Changement ici

            var lastColUsed = worksheet.LastColumnUsed();
            if (lastColUsed == null)
            {
                await LogAndSend("❌ Erreur: La feuille de calcul ne contient aucune colonne utilisée.", cancellationToken);
                throw new InvalidOperationException("La feuille de calcul ne contient aucune colonne utilisée.");
            }
            int lastColumn = lastColUsed.ColumnNumber();
            LogOnly($"Dernière colonne utilisée détectée : {lastColumn}"); // Changement ici

            // Création du dossier de sortie si inexistant
            if (!Directory.Exists(outputDir))
            {
                await LogAndSend($"Création du dossier de sortie : {outputDir}", cancellationToken);
                Directory.CreateDirectory(outputDir);
            }
            else
            {
                await LogAndSend($"Le dossier de sortie existe déjà : {outputDir}", cancellationToken);
            }

            // --- NOUVELLE LOGIQUE DE DÉLIMITATION DES BLOCS BASÉE SUR LE CYCLE DES DATES ---
            LogOnly("Étape 1: Recherche des lignes contenant des dates..."); // Changement ici
            List<int> dateLines = new List<int>();
            for (int row = 1; row <= lastRow; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var cell = worksheet.Cell(row, 1);
                string text = cell.GetString();
                if (DateTime.TryParse(text, out _))
                {
                    dateLines.Add(row);
                }
            }

            if (dateLines.Count == 0)
            {
                await LogAndSend("Aucune date trouvée dans la feuille de calcul. Impossible de délimiter les blocs partenaires.", cancellationToken);
                return;
            }
            LogOnly($"{dateLines.Count} dates trouvées dans la colonne A."); // Changement ici


            // --- AJOUT : Trouver la date la plus basse et la plus haute du fichier entier ---
            LogOnly("Étape 2: Détermination de la plage de dates globale du fichier..."); // Changement ici
            DateTime overallMinDate = DateTime.MaxValue;
            DateTime overallMaxDate = DateTime.MinValue;

            foreach (int dateRow in dateLines)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (DateTime.TryParse(worksheet.Cell(dateRow, 1).GetString(), out DateTime currentParsedDate))
                {
                    if (currentParsedDate < overallMinDate)
                    {
                        overallMinDate = currentParsedDate;
                    }
                    if (currentParsedDate > overallMaxDate)
                    {
                        overallMaxDate = currentParsedDate;
                    }
                }
            }

            string dateStrmin = (overallMinDate != DateTime.MaxValue) ? overallMinDate.ToString("dd.MM.yyyy") : "DateMinInconnue";
            string dateStrmax = (overallMaxDate != DateTime.MinValue) ? overallMaxDate.ToString("dd.MM.yyyy") : "DateMaxInconnue";

            LogOnly($"Plage de dates globale identifiée : du {dateStrmin} au {dateStrmax}"); // Changement ici


            LogOnly("Étape 3: Délimitation des blocs partenaires basée sur le cycle des dates..."); // Changement ici
            List<(int startRow, int endRow)> partnerBlocks = new List<(int startRow, int endRow)>();
            DateTime? previousDate = null;
            int? currentBlockStartRow = null;

            currentBlockStartRow = dateLines[0] - 1;
            LogOnly($"Début de la délimitation des blocs. Premier bloc potentiel commence à la ligne {currentBlockStartRow}"); // Changement ici


            for (int i = 0; i < dateLines.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int currentRowDateIndex = dateLines[i];
                DateTime currentDate;

                if (!DateTime.TryParse(worksheet.Cell(currentRowDateIndex, 1).GetString(), out currentDate))
                {
                    LogOnly($"Avertissement: Impossible de parser la date à la ligne {currentRowDateIndex}. La ligne sera ignorée pour la délimitation."); // Changement ici
                    continue;
                }

                if (previousDate.HasValue && currentDate <= previousDate.Value)
                {
                    if (currentBlockStartRow.HasValue)
                    {
                        LogOnly($"Bloc délimité : lignes {currentBlockStartRow.Value} à {currentRowDateIndex - 2}"); // Changement ici
                        partnerBlocks.Add((currentBlockStartRow.Value, currentRowDateIndex - 2));
                    }
                    currentBlockStartRow = currentRowDateIndex - 1;
                    LogOnly($"Nouveau cycle de date détecté. Nouveau bloc commence à la ligne {currentBlockStartRow}"); // Changement ici
                }

                previousDate = currentDate;
            }

            if (currentBlockStartRow.HasValue)
            {
                partnerBlocks.Add((currentBlockStartRow.Value, lastRow));
                LogOnly($"Dernier bloc délimité : lignes {currentBlockStartRow.Value} à {lastRow}"); // Changement ici
            }

            int totalPartners = partnerBlocks.Count;
            await LogAndSend($"Total de {totalPartners} blocs partenaires identifiés.", cancellationToken); // Reste important pour l'utilisateur


            if (totalPartners == 0)
            {
                await LogAndSend("Aucun bloc partenaire identifiable trouvé dans le fichier Excel selon la logique des dates.", cancellationToken);
                return;
            }

            // Ajustement des index et du compte pour la boucle
            if (startIndex < 0) startIndex = 0;
            if (startIndex >= totalPartners) startIndex = totalPartners - 1;
            if (count < 1) count = 1;
            if (count > totalPartners - startIndex) count = totalPartners - startIndex;
            await LogAndSend($"Traitement prévu pour {count} blocs partenaires, à partir de l'index {startIndex}.", cancellationToken); // Reste important pour l'utilisateur

            // Initialisation de la progression
            await _hubContext.Clients.All.SendAsync("ReceiveProgress", new
            {
                Current = 0,
                Total = count,
                Percentage = 0,
                Message = "Début de la génération des fichiers partenaires." // Message initial pour l'utilisateur
            }, cancellationToken);

            LogOnly($"--- Début de la génération des fichiers Excel par partenaire ---"); // Changement ici

            // 3. Boucle principale pour traiter chaque bloc identifié
            for (int i = startIndex; i < startIndex + count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var (blockStartRow, blockEndRow) = partnerBlocks[i];

                try
                {
                    string partnerName = worksheet.Row(blockStartRow).Cell(1).GetString().Trim();
                    LogOnly($"Traitement du partenaire '{partnerName}' (Bloc {i + 1}/{totalPartners} - Lignes {blockStartRow}-{blockEndRow})"); // Changement ici

                    DateTime blockDate = DateTime.MinValue;
                    var cellForBlockDate = worksheet.Cell(blockStartRow + 1, 1);
                    if (DateTime.TryParse(cellForBlockDate.GetString(), out DateTime parsedDate))
                    {
                        blockDate = parsedDate;
                        LogOnly($"  - Date de début du bloc détectée : {blockDate:dd.MM.yyyy}"); // Changement ici
                    }
                    else
                    {
                        LogOnly($"  - ⚠️ Aucune date valide trouvée à la ligne {blockStartRow + 1} pour le bloc '{partnerName}'."); // Changement ici
                    }


                    LogOnly($"  - Ouverture du fichier template : {templatePath}"); // Changement ici
                    using var templateWb = new XLWorkbook(templatePath);
                    var templateWs = templateWb.Worksheet(1);
                    LogOnly("  - Template ouvert avec succès."); // Changement ici

                    int currentTargetRow = 3;
                    LogOnly($"  - Copie des lignes du bloc ({blockStartRow} à {blockEndRow}) vers le template (à partir de la ligne {currentTargetRow})..."); // Changement ici
                    for (int r = blockStartRow; r <= blockEndRow; r++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var sourceRow = worksheet.Row(r);
                        var targetRow = templateWs.Row(currentTargetRow);

                        for (int c = 1; c <= lastColumn; c++)
                        {
                            var sourceCell = sourceRow.Cell(c);
                            var targetCell = targetRow.Cell(c);
                            targetCell.Value = sourceCell.Value;
                            targetCell.Style = sourceCell.Style;
                        }
                        currentTargetRow++;
                    }
                    LogOnly($"  - {blockEndRow - blockStartRow + 1} lignes copiées dans le template."); // Changement ici


                    int templateLastRow = templateWs.LastRowUsed()?.RowNumber() ?? 0;
                    if (templateLastRow >= currentTargetRow)
                    {
                        LogOnly($"  - Suppression des lignes excédentaires du template (lignes {currentTargetRow} à {templateLastRow})..."); // Changement ici
                        for (int rowToDelete = templateLastRow; rowToDelete >= currentTargetRow; rowToDelete--)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            templateWs.Row(rowToDelete).Delete();
                        }
                        LogOnly("  - Lignes excédentaires supprimées."); // Changement ici
                    }
                    else
                    {
                        LogOnly("  - Aucune ligne excédentaire à supprimer dans le template."); // Changement ici
                    }


                        //AJOUT DES FEUILLES SUPPLÉMENTAIRES ICI
                    LogOnly($"  - Ajout des feuilles supplémentaires pour '{partnerName}'..."); // Changement ici
                    var feuillesAScanner = new List<string> { "Activité nette à J", "J+1", "Regul" };
                    await AddSupplementarySheetsAsync(worksheet.Workbook, templateWb, partnerName, feuillesAScanner, cancellationToken);
                    LogOnly($"  - Traitement des feuilles supplémentaires terminé pour '{partnerName}'."); // Changement ici

                     // Création du nom de fichier de sortie
                    string safePartnerName = string.Concat(partnerName.Split(Path.GetInvalidFileNameChars()));
                    string dateRange = (dateStrmin == dateStrmax) ? dateStrmin : $"{dateStrmin} au {dateStrmax}";
                    string outputFileName = $"COMPTE SUPPORT {safePartnerName} du {dateRange}.xlsx";
                    string outputPath = Path.Combine(outputDir, outputFileName);

                    LogOnly($"  - Sauvegarde du fichier : {outputFileName} dans {outputDir}");
                    templateWb.SaveAs(outputPath);
                    await LogAndSend($"✅ Fichier '{outputFileName}' généré avec succès.", cancellationToken);
                }

                catch (OperationCanceledException)
                {
                    await LogAndSend("❌ Génération annulée par l'utilisateur.", CancellationToken.None);
                    throw;
                }
                catch (Exception ex)
                {
                    string blockTitle = (blockStartRow > 0 && blockStartRow <= lastRow) ? worksheet.Row(blockStartRow).Cell(1).GetString() : "Inconnu";
                    await LogAndSend($"❌ Erreur inattendue lors de la génération du fichier pour le bloc '{blockTitle}' (lignes {blockStartRow}-{blockEndRow}) : {ex.Message}", CancellationToken.None); // Message important pour l'utilisateur
                    LogOnly($"(Détails erreur: {ex.StackTrace})"); // StackTrace seulement en console
                }

                int currentProcessed = i - startIndex + 1;
                double percentage = (double)currentProcessed / count * 100;

                // --- MODIFICATION ICI : Utilisation de ProgressUpdate pour l'envoi ---
                await _hubContext.Clients.All.SendAsync("ReceiveProgress", new ProgressUpdate // <-- UTILISEZ ProgressUpdate ICI
                {
                    Current = currentProcessed,
                    Total = count,
                    Percentage = (int)percentage,
                    Message = $"Progression : {(int)percentage}% - Fichier {currentProcessed} sur {count} généré."
                }, cancellationToken);

                //await LogAndSend($"Progression : {(int)percentage}% - Fichier {currentProcessed} sur {count} généré.", cancellationToken);

            }

            // Message final de progression (100%)
            await _hubContext.Clients.All.SendAsync("ReceiveProgress", new
            {
                Current = count,
                Total = count,
                Percentage = 100,
                Message = "Génération terminée." // Message final pour l'utilisateur
            }, cancellationToken);
            await LogAndSend("--- Processus de génération des fichiers partenaires terminé ---", cancellationToken); // Final summary
        }

        public Task AddSupplementarySheetsAsync(
                XLWorkbook sourceWorkbook,
                XLWorkbook partnerWorkbook,
                string partnerName,
                List<string> feuillesAScanner,
                CancellationToken cancellationToken = default)
        {
            LogOnly($"    Début de l'ajout des feuilles supplémentaires pour '{partnerName}'."); // Changement ici
            foreach (var feuilleName in feuillesAScanner)
            {
                cancellationToken.ThrowIfCancellationRequested();

                LogOnly($"      - Traitement de la feuille '{feuilleName}'..."); // Changement ici
                var sourceSheet = sourceWorkbook.Worksheets
                    .FirstOrDefault(ws => string.Equals(ws.Name.Trim(), feuilleName.Trim(), StringComparison.OrdinalIgnoreCase));
                if (sourceSheet == null)
                {
                    LogOnly($"      ❌ Feuille '{feuilleName}' introuvable dans le classeur source. Ignorée."); // Changement ici
                    continue;
                }
                LogOnly($"      Feuille source '{feuilleName}' trouvée."); // Changement ici

                var lastRowUsed = sourceSheet.LastRowUsed();
                if (lastRowUsed == null)
                {
                    LogOnly($"      ❌ Feuille '{feuilleName}' est vide. Ignorée."); // Changement ici
                    continue;
                }
                int lastRow = lastRowUsed.RowNumber();
                LogOnly($"      Feuille '{feuilleName}' a {lastRow} lignes utilisées."); // Changement ici


                int lastCol = sourceSheet.LastColumnUsed()?.ColumnNumber() ?? 0;
                var headerRow = sourceSheet.Row(1);

                int colDistributeur = -1;
                for (int col = 1; col <= lastCol; col++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var cell = headerRow.Cell(col);
                    if (cell != null && cell.GetString().Trim().Equals("Distributeur", StringComparison.OrdinalIgnoreCase))
                    {
                        colDistributeur = col;
                        break;
                    }
                }

                if (colDistributeur == -1)
                {
                    LogOnly($"      ❌ Colonne 'Distributeur' introuvable dans la feuille '{feuilleName}'. Ignorée."); // Changement ici
                    continue;
                }
                LogOnly($"      Colonne 'Distributeur' trouvée à l'index {colDistributeur} dans '{feuilleName}'."); // Changement ici

                var matchedRows = new List<IXLRow>();
                LogOnly($"      Recherche des lignes pour '{partnerName}' dans la feuille '{feuilleName}'..."); // Changement ici
                for (int row = 2; row <= lastRow; row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var cell = sourceSheet.Cell(row, colDistributeur);
                    if (cell?.GetString().Trim().Equals(partnerName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        matchedRows.Add(sourceSheet.Row(row));
                    }
                }

                if (matchedRows.Count == 0)
                {
                    LogOnly($"      Aucune ligne trouvée pour le partenaire '{partnerName}' dans la feuille '{feuilleName}'."); // Changement ici
                    continue;
                }
                LogOnly($"      {matchedRows.Count} lignes trouvées pour '{partnerName}' dans la feuille '{feuilleName}'."); // Changement ici


                // Vérifier si la feuille existe déjà dans le classeur partenaire
                if (partnerWorkbook.Worksheets.Any(ws => ws.Name == feuilleName))
                {
                    if (feuilleName.Equals("Activité nette à J", StringComparison.OrdinalIgnoreCase))
                    {
                        var sheetToDelete = partnerWorkbook.Worksheet(feuilleName);
                        partnerWorkbook.Worksheets.Delete(sheetToDelete.Name);
                        LogOnly($"      La feuille existante '{feuilleName}' a été supprimée pour être remplacée pour '{partnerName}'."); // Changement ici
                    }
                    else
                    {
                        LogOnly($"      La feuille '{feuilleName}' existe déjà pour '{partnerName}', elle ne sera pas ajoutée pour éviter les doublons."); // Changement ici
                        continue;
                    }
                }

                LogOnly($"      Ajout de la nouvelle feuille '{feuilleName}' au classeur partenaire..."); // Changement ici
                var newSheet = partnerWorkbook.AddWorksheet(feuilleName);

                LogOnly("      Copie de l'en-tête de la feuille..."); // Changement ici
                for (int col = 1; col <= lastCol; col++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var sourceCell = headerRow.Cell(col);
                    var destCell = newSheet.Cell(1, col);
                    destCell.Value = sourceCell.Value;
                    destCell.Style = sourceCell.Style;
                }
                LogOnly("      En-tête copié."); // Changement ici

                int currentRow = 2;
                LogOnly($"      Copie des {matchedRows.Count} lignes correspondantes..."); // Changement ici
                foreach (var row in matchedRows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int col = 1; col <= lastCol; col++)
                    {
                        var sourceCell = row.Cell(col);
                        var destCell = newSheet.Cell(currentRow, col);

                        destCell.Value = sourceCell.Value;
                        destCell.Style = sourceCell.Style;

                        if (sourceCell.DataType == XLDataType.DateTime && sourceCell.GetDateTime() != DateTime.MinValue)
                        {
                            destCell.Value = sourceCell.GetDateTime();
                            destCell.Style.DateFormat.Format = "dd/MM/yyyy";
                        }
                    }
                    currentRow++;
                }
                LogOnly("      Lignes de données copiées."); // Changement ici

                newSheet.Columns().AdjustToContents();
                LogOnly($"      ✅ Feuille '{feuilleName}' ajoutée et ajustée pour le partenaire '{partnerName}' avec {matchedRows.Count} lignes de données."); // Changement ici
            }
            LogOnly("    Fin de l'ajout des feuilles supplémentaires."); // Changement ici
            return Task.CompletedTask;
        }
    }
}