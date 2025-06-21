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
            LogOnly($"Dernière ligne utilisée détectée : {lastRow}");

            var lastColUsed = worksheet.LastColumnUsed();
            if (lastColUsed == null)
            {
                await LogAndSend("❌ Erreur: La feuille de calcul ne contient aucune colonne utilisée.", cancellationToken);
                throw new InvalidOperationException("La feuille de calcul ne contient aucune colonne utilisée.");
            }
            int lastColumn = lastColUsed.ColumnNumber();
            LogOnly($"Dernière colonne utilisée détectée : {lastColumn}");

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

            // Étape 1: Recherche des lignes contenant des dates
            LogOnly("Étape 1: Recherche des lignes contenant des dates...");
            List<int> dateLines = new List<int>();
            Dictionary<int, string> colorInfoCache = new Dictionary<int, string>();

            // Table de correspondance pour les couleurs indexées (palette Excel par défaut)
            Dictionary<int, string> indexedColorMap = new Dictionary<int, string>
        {
            { 64, "#FFFFFF" }, // Index 64 correspond souvent à une couleur blanche ou par défaut
            // Ajoutez d'autres indices si nécessaire, selon la palette Excel
        };

            // Table de correspondance approximative pour les couleurs de thème
            Dictionary<XLThemeColor, string> themeColorMap = new Dictionary<XLThemeColor, string>
        {
            { XLThemeColor.Accent4, "#4BACC6" }, // Approximation pour Accent4
            { XLThemeColor.Background1, "#FFFFFF" }, // Approximation pour Background1
            // Ajoutez d'autres couleurs de thème si nécessaire
        };

            for (int row = 1; row <= lastRow; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var cell = worksheet.Cell(row, 1);
                string text = cell.GetString();

                // Récupère la couleur de fond
                var bgColor = cell.Style.Fill.BackgroundColor;
                string colorInfo;

                if (bgColor.ColorType == XLColorType.Color)
                {
                    var color = bgColor.Color; // Couleur RGB directe
                    colorInfo = $"#{color.R:X2}{color.G:X2}{color.B:X2}"; // Format hexadécimal
                }
                else if (bgColor.ColorType == XLColorType.Theme)
                {
                    try
                    {
                        var themeColor = bgColor.ThemeColor;
                        var tint = bgColor.ThemeTint;
                        if (themeColorMap.ContainsKey(themeColor))
                        {
                            colorInfo = themeColorMap[themeColor];
                            if (tint != 0)
                            {
                                colorInfo += $", Tint: {tint}";
                            }
                        }
                        else
                        {
                            colorInfo = $"Theme: {bgColor.ToString()}"; // Fallback
                        }
                    }
                    catch
                    {
                        colorInfo = $"Theme: {bgColor.ToString()}"; // Fallback en cas d'erreur
                    }
                }
                else if (bgColor.ColorType == XLColorType.Indexed)
                {
                    int colorIndex = bgColor.Indexed;
                    colorInfo = indexedColorMap.ContainsKey(colorIndex) ? indexedColorMap[colorIndex] : $"Color Index: {colorIndex}";
                }
                else
                {
                    colorInfo = bgColor.ToString(); // Couleurs nommées, transparent ou autres
                }

                colorInfoCache[row] = colorInfo; // Cache pour réutilisation dans la délimitation des blocs

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
            LogOnly($"{dateLines.Count} dates trouvées dans la colonne A.");

            // Étape 2: Détermination de la plage de dates globale du fichier
            LogOnly("Étape 2: Détermination de la plage de dates globale du fichier...");
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
            LogOnly($"Plage de dates globale identifiée : du {dateStrmin} au {dateStrmax}");

            // Étape 3: Nouvelle logique de délimitation des blocs
            LogOnly("Étape 3: Délimitation des blocs partenaires basée sur la nouvelle logique...");
            List<(int startRow, int endRow)> partnerBlocks = new List<(int startRow, int endRow)>();
            int? currentBlockStartRow = null;

            // Le premier bloc commence à la ligne qui précède la première date
            if (dateLines.Count > 0)
            {
                currentBlockStartRow = Math.Max(1, dateLines[0] - 1);
                LogOnly($"Premier bloc commence à la ligne {currentBlockStartRow}");
            }

            for (int row = 1; row <= lastRow; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var cell = worksheet.Cell(row, 1);
                string text = cell.GetString();
                bool isDate = DateTime.TryParse(text, out _);
                bool isColorIndex64 = cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Indexed && cell.Style.Fill.BackgroundColor.Indexed == 64;

                // Vérifier si la ligne déclenche un nouveau bloc
                if (!isDate && !isColorIndex64 && currentBlockStartRow.HasValue && row > currentBlockStartRow.Value)
                {
                    // Fermer le bloc précédent
                    partnerBlocks.Add((currentBlockStartRow.Value, row - 1));
                    LogOnly($"Bloc délimité : lignes {currentBlockStartRow.Value} à {row - 1}");
                    currentBlockStartRow = row; // Nouveau bloc commence à la ligne courante
                    LogOnly($"Nouveau bloc commence à la ligne {currentBlockStartRow}");
                }
            }

            // Ajouter le dernier bloc
            if (currentBlockStartRow.HasValue)
            {
                partnerBlocks.Add((currentBlockStartRow.Value, lastRow));
                LogOnly($"Dernier bloc délimité : lignes {currentBlockStartRow.Value} à {lastRow}");
            }

            int totalPartners = partnerBlocks.Count;
            await LogAndSend($"Total de {totalPartners} blocs partenaires identifiés.", cancellationToken);

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
            await LogAndSend($"Traitement prévu pour {count} blocs partenaires, à partir de l'index {startIndex}.", cancellationToken);

            // Initialisation de la progression
            await _hubContext.Clients.All.SendAsync("ReceiveProgress", new
            {
                Current = 0,
                Total = count,
                Percentage = 0,
                Message = "Début de la génération des fichiers partenaires."
            }, cancellationToken);

            LogOnly($"--- Début de la génération des fichiers Excel par partenaire ---");

            // Boucle principale pour traiter chaque bloc identifié
            for (int i = startIndex; i < startIndex + count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var (blockStartRow, blockEndRow) = partnerBlocks[i];

                try
                {
                    string partnerName = worksheet.Row(blockStartRow).Cell(1).GetString().Trim();
                    LogOnly($"Traitement du partenaire '{partnerName}' (Bloc {i + 1}/{totalPartners} - Lignes {blockStartRow}-{blockEndRow})");

                    DateTime blockDate = DateTime.MinValue;
                    var cellForBlockDate = worksheet.Cell(blockStartRow + 1, 1);
                    if (DateTime.TryParse(cellForBlockDate.GetString(), out DateTime parsedDate))
                    {
                        blockDate = parsedDate;
                        LogOnly($"  - Date de début du bloc détectée : {blockDate:dd.MM.yyyy}");
                    }
                    else
                    {
                        LogOnly($"  - ⚠️ Aucune date valide trouvée à la ligne {blockStartRow + 1} pour le bloc '{partnerName}'.");
                    }

                    LogOnly($"  - Ouverture du fichier template : {templatePath}");
                    using var templateWb = new XLWorkbook(templatePath);
                    var templateWs = templateWb.Worksheet(1);
                    LogOnly("  - Template ouvert avec succès.");

                    int currentTargetRow = 3;
                    LogOnly($"  - Copie des lignes du bloc ({blockStartRow} à {blockEndRow}) vers le template (à partir de la ligne {currentTargetRow})...");
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
                    LogOnly($"  - {blockEndRow - blockStartRow + 1} lignes copiées dans le template.");

                    // Ajuster automatiquement la largeur des colonnes de la feuille principale
                    templateWs.Columns().AdjustToContents();
                    foreach (var column in templateWs.ColumnsUsed())
                    {
                        column.Width += 8;
                    }
                    LogOnly("  - Colonnes ajustées automatiquement aux contenus.");

                    // Appliquer la police "Calibri", taille 10 à toute la feuille
                    templateWs.Style.Font.FontName = "Calibri";
                    templateWs.Style.Font.FontSize = 10;

                    int templateLastRow = templateWs.LastRowUsed()?.RowNumber() ?? 0;
                    if (templateLastRow >= currentTargetRow)
                    {
                        LogOnly($"  - Suppression des lignes excédentaires du template (lignes {currentTargetRow} à {templateLastRow})...");
                        for (int rowToDelete = templateLastRow; rowToDelete >= currentTargetRow; rowToDelete--)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            templateWs.Row(rowToDelete).Delete();
                        }
                        LogOnly("  - Lignes excédentaires supprimées.");
                    }
                    else
                    {
                        LogOnly("  - Aucune ligne excédentaire à supprimer dans le template.");
                    }

                    // Ajout des feuilles supplémentaires
                    LogOnly($"  - Ajout des feuilles supplémentaires pour '{partnerName}'...");
                    var feuillesAScanner = new List<string> { "Activité nette à J", "J+1", "Regul", "Distributions" };
                    await AddSupplementarySheetsAsync(worksheet.Workbook, templateWb, partnerName, feuillesAScanner, cancellationToken);
                    LogOnly($"  - Traitement des feuilles supplémentaires terminé pour '{partnerName}'.");

                    // Création du nom de fichier de sortie
                    string safePartnerName = string.Concat(partnerName.Split(Path.GetInvalidFileNameChars()));
                    string dateRange = (dateStrmin == dateStrmax) ? dateStrmin : $"{dateStrmin} au {dateStrmax}";
                    string outputFileName = $"COMPTE SUPPORT {safePartnerName} du {dateRange}.xlsx";
                    string outputPath = Path.Combine(outputDir, outputFileName);

                    LogOnly($"  - Sauvegarde du fichier : {outputFileName} dans {outputDir}");
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
                    await LogAndSend($"❌ Erreur inattendue lors de la génération du fichier pour le bloc '{blockTitle}' (lignes {blockStartRow}-{blockEndRow}) : {ex.Message}", CancellationToken.None);
                    LogOnly($"(Détails erreur: {ex.StackTrace})");
                }

                int currentProcessed = i - startIndex + 1;
                double percentage = (double)currentProcessed / count * 100;

                await _hubContext.Clients.All.SendAsync("ReceiveProgress", new
                {
                    Current = currentProcessed,
                    Total = count,
                    Percentage = (int)percentage,
                    Message = $"Progression : {(int)percentage}% - Fichier {currentProcessed} sur {count} généré."
                }, cancellationToken);
            }

            // Message final de progression (100%)
            await _hubContext.Clients.All.SendAsync("ReceiveProgress", new
            {
                Current = count,
                Total = count,
                Percentage = 100,
                Message = "Génération terminée."
            }, cancellationToken);
            await LogAndSend("--- Processus de génération des fichiers partenaires terminé ---", cancellationToken);
        }

        public Task AddSupplementarySheetsAsync(
            XLWorkbook sourceWorkbook,
            XLWorkbook partnerWorkbook,
            string partnerName,
            List<string> feuillesAScanner,
            CancellationToken cancellationToken = default)
        {
            LogOnly($"    Début de l'ajout des feuilles supplémentaires pour '{partnerName}'.");

            foreach (var feuilleName in feuillesAScanner)
            {
                cancellationToken.ThrowIfCancellationRequested();

                LogOnly($"      - Traitement de la feuille '{feuilleName}'...");
                var sourceSheet = sourceWorkbook.Worksheets
                    .FirstOrDefault(ws => string.Equals(ws.Name.Trim(), feuilleName.Trim(), StringComparison.OrdinalIgnoreCase));

                if (sourceSheet == null)
                {
                    LogOnly($"      ❌ Feuille '{feuilleName}' introuvable dans le classeur source. Ignorée.");
                    continue;
                }

                LogOnly($"      Feuille source '{feuilleName}' trouvée.");
                var lastRowUsed = sourceSheet.LastRowUsed();

                if (lastRowUsed == null)
                {
                    LogOnly($"      ❌ Feuille '{feuilleName}' est vide. Ignorée.");
                    continue;
                }

                int lastRow = lastRowUsed.RowNumber();
                LogOnly($"      Feuille '{feuilleName}' a {lastRow} lignes utilisées.");

                int lastCol = sourceSheet.LastColumnUsed()?.ColumnNumber() ?? 0;
                var headerRow = sourceSheet.Row(1);

                int colParentReseau = -1;
                for (int col = 1; col <= lastCol; col++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var cell = headerRow.Cell(col);
                    if (cell != null && cell.GetString().Trim().Equals("Parent Réseau", StringComparison.OrdinalIgnoreCase))
                    {
                        colParentReseau = col;
                        break;
                    }
                }

                if (colParentReseau == -1)
                {
                    LogOnly($"      ❌ Colonne 'Parent Réseau' introuvable dans la feuille '{feuilleName}'. Ignorée.");
                    continue;
                }

                LogOnly($"      Colonne 'Parent Réseau' trouvée à l'index {colParentReseau} dans '{feuilleName}'.");

                var matchedRows = new List<IXLRow>();
                LogOnly($"      Recherche des lignes pour '{partnerName}' dans la feuille '{feuilleName}'...");

                for (int row = 2; row <= lastRow; row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var cell = sourceSheet.Cell(row, colParentReseau);
                    if (cell?.GetString().Trim().Equals(partnerName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        matchedRows.Add(sourceSheet.Row(row));
                    }
                }

                if (matchedRows.Count == 0)
                {
                    LogOnly($"      Aucune ligne trouvée pour le partenaire '{partnerName}' dans la feuille '{feuilleName}'.");
                    continue;
                }

                LogOnly($"      {matchedRows.Count} lignes trouvées pour '{partnerName}' dans la feuille '{feuilleName}'.");

                if (partnerWorkbook.Worksheets.Any(ws => ws.Name == feuilleName))
                {
                    if (feuilleName.Equals("Activité nette à J", StringComparison.OrdinalIgnoreCase))
                    {
                        var sheetToDelete = partnerWorkbook.Worksheet(feuilleName);
                        partnerWorkbook.Worksheets.Delete(sheetToDelete.Name);
                        LogOnly($"      La feuille existante '{feuilleName}' a été supprimée pour être remplacée pour '{partnerName}'.");
                    }
                    else
                    {
                        LogOnly($"      La feuille '{feuilleName}' existe déjà pour '{partnerName}', elle ne sera pas ajoutée pour éviter les doublons.");
                        continue;
                    }
                }

                LogOnly($"      Ajout de la nouvelle feuille '{feuilleName}' au classeur partenaire...");
                var newSheet = partnerWorkbook.AddWorksheet(feuilleName);

                LogOnly("      Copie de l'en-tête de la feuille...");
                for (int col = 1; col <= lastCol; col++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var sourceCell = headerRow.Cell(col);
                    var destCell = newSheet.Cell(1, col);
                    destCell.Value = sourceCell.Value;
                    destCell.Style = sourceCell.Style;
                }

                LogOnly("      En-tête copié.");
                int currentRow = 2;

                LogOnly($"      Copie des {matchedRows.Count} lignes correspondantes...");
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

                LogOnly("      Lignes de données copiées.");
                newSheet.Columns().AdjustToContents();
                LogOnly("  - Toutes les colonnes ont été ajustées automatiquement.");

                newSheet.Style.Font.FontName = "Calibri";
                newSheet.Style.Font.FontSize = 10;

                // Si on est dans la feuille "Distributions", augmenter encore la largeur de toutes les colonnes
                if (feuilleName.Equals("Distributions", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var column in newSheet.ColumnsUsed())
                    {
                        column.Width += 8;  // Augmente la largeur par 8 unités (à ajuster selon besoin)
                    }
                }

                LogOnly($"      ✅ Feuille '{feuilleName}' ajoutée et ajustée pour le partenaire '{partnerName}' avec {matchedRows.Count} lignes de données.");
            }

            LogOnly("    Fin de l'ajout des feuilles supplémentaires.");
            return Task.CompletedTask;
        }

    }
}