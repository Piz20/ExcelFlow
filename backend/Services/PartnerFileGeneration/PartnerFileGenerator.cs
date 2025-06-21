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
    await LogAndSend("🚀 Lancement du processus de génération des fichiers Excel pour les partenaires...", cancellationToken);

    // Vérifications initiales de la feuille
    var lastRowUsed = worksheet.LastRowUsed();
    if (lastRowUsed == null)
    {
        await LogAndSend("❌ Impossible de continuer : la feuille Excel est vide (aucune ligne détectée).", cancellationToken);
        throw new InvalidOperationException("La feuille de calcul ne contient aucune ligne utilisée.");
    }
    int lastRow = lastRowUsed.RowNumber();
    LogOnly($"Dernière ligne utilisée détectée : {lastRow}");

    var lastColUsed = worksheet.LastColumnUsed();
    if (lastColUsed == null)
    {
        await LogAndSend("❌ Impossible de continuer : aucune colonne détectée dans la feuille Excel.", cancellationToken);
        throw new InvalidOperationException("La feuille de calcul ne contient aucune colonne utilisée.");
    }
    int lastColumn = lastColUsed.ColumnNumber();
    LogOnly($"Dernière colonne utilisée détectée : {lastColumn}");

    await LogAndSend("✅ Feuille Excel analysée : lignes et colonnes détectées avec succès.", cancellationToken);

    // Création du dossier de sortie si inexistant
    if (!Directory.Exists(outputDir))
    {
        await LogAndSend($"📁 Création du dossier de sortie : {outputDir}", cancellationToken);
        Directory.CreateDirectory(outputDir);
    }
    else
    {
        await LogAndSend($"📁 Dossier de sortie détecté : {outputDir}", cancellationToken);
    }

    // Étape 1: Recherche des lignes contenant des dates
    await LogAndSend("🔎 Recherche des lignes contenant des dates dans la colonne A...", cancellationToken);
    List<int> dateLines = new();
    Dictionary<int, string> colorInfoCache = new();
    Dictionary<int, string> indexedColorMap = new() { { 64, "#FFFFFF" } };
    Dictionary<XLThemeColor, string> themeColorMap = new()
    {
        { XLThemeColor.Accent4, "#4BACC6" },
        { XLThemeColor.Background1, "#FFFFFF" }
    };

    for (int row = 1; row <= lastRow; row++)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var cell = worksheet.Cell(row, 1);
        string text = cell.GetString();
        var bgColor = cell.Style.Fill.BackgroundColor;
        string colorInfo;

        if (bgColor.ColorType == XLColorType.Color)
        {
            var color = bgColor.Color;
            colorInfo = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }
        else if (bgColor.ColorType == XLColorType.Theme)
        {
            try
            {
                var themeColor = bgColor.ThemeColor;
                var tint = bgColor.ThemeTint;
                colorInfo = themeColorMap.ContainsKey(themeColor)
                    ? themeColorMap[themeColor] + (tint != 0 ? $", Tint: {tint}" : "")
                    : $"Theme: {bgColor}";
            }
            catch
            {
                colorInfo = $"Theme: {bgColor}";
            }
        }
        else if (bgColor.ColorType == XLColorType.Indexed)
        {
            int colorIndex = bgColor.Indexed;
            colorInfo = indexedColorMap.ContainsKey(colorIndex) ? indexedColorMap[colorIndex] : $"Color Index: {colorIndex}";
        }
        else
        {
            colorInfo = bgColor.ToString();
        }

        colorInfoCache[row] = colorInfo;

        if (DateTime.TryParse(text, out _))
        {
            dateLines.Add(row);
        }
    }

    if (dateLines.Count == 0)
    {
        await LogAndSend("⚠️ Aucune ligne contenant une date n'a été trouvée. Impossible de détecter les blocs partenaires.", cancellationToken);
        return;
    }

    await LogAndSend($"📅 {dateLines.Count} ligne(s) contenant des dates détectée(s).", cancellationToken);

    // Étape 2: Détermination de la plage de dates
    DateTime overallMinDate = DateTime.MaxValue;
    DateTime overallMaxDate = DateTime.MinValue;

    foreach (int dateRow in dateLines)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (DateTime.TryParse(worksheet.Cell(dateRow, 1).GetString(), out DateTime currentParsedDate))
        {
            if (currentParsedDate < overallMinDate) overallMinDate = currentParsedDate;
            if (currentParsedDate > overallMaxDate) overallMaxDate = currentParsedDate;
        }
    }

    string dateStrmin = (overallMinDate != DateTime.MaxValue) ? overallMinDate.ToString("dd.MM.yyyy") : "DateMinInconnue";
    string dateStrmax = (overallMaxDate != DateTime.MinValue) ? overallMaxDate.ToString("dd.MM.yyyy") : "DateMaxInconnue";

    await LogAndSend($"📆 Plage de dates détectée : du {dateStrmin} au {dateStrmax}.", cancellationToken);

    // Étape 3: Délimitation des blocs partenaires
    await LogAndSend("📦 Délimitation des blocs partenaires à partir des lignes analysées...", cancellationToken);
    List<(int startRow, int endRow)> partnerBlocks = new();
    int? currentBlockStartRow = null;

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

        if (!isDate && !isColorIndex64 && currentBlockStartRow.HasValue && row > currentBlockStartRow.Value)
        {
            partnerBlocks.Add((currentBlockStartRow.Value, row - 1));
            LogOnly($"Bloc délimité : lignes {currentBlockStartRow.Value} à {row - 1}");
            currentBlockStartRow = row;
        }
    }

    if (currentBlockStartRow.HasValue)
    {
        partnerBlocks.Add((currentBlockStartRow.Value, lastRow));
        LogOnly($"Dernier bloc délimité : lignes {currentBlockStartRow.Value} à {lastRow}");
    }

    int totalPartners = partnerBlocks.Count;
    await LogAndSend($"✅ {totalPartners} bloc(s) partenaire(s) identifié(s).", cancellationToken);

    if (totalPartners == 0)
    {
        await LogAndSend("❌ Aucun bloc partenaire identifiable trouvé.", cancellationToken);
        return;
    }

    if (startIndex < 0) startIndex = 0;
    if (startIndex >= totalPartners) startIndex = totalPartners - 1;
    if (count < 1) count = 1;
    if (count > totalPartners - startIndex) count = totalPartners - startIndex;

    await LogAndSend($"📊 {count} bloc(s) seront traités à partir de l’index {startIndex}.", cancellationToken);

    await _hubContext.Clients.All.SendAsync("ReceiveProgress", new
    {
        Current = 0,
        Total = count,
        Percentage = 0,
        Message = "🔄 Début du traitement des partenaires..."
    }, cancellationToken);

    LogOnly($"--- Début de la génération des fichiers Excel par partenaire ---");

    for (int i = startIndex; i < startIndex + count; i++)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var (blockStartRow, blockEndRow) = partnerBlocks[i];

        try
        {
            string partnerName = worksheet.Row(blockStartRow).Cell(1).GetString().Trim();
            await LogAndSend($"📂 Traitement du partenaire {i + 1}/{totalPartners} : '{partnerName}'...", cancellationToken);

            DateTime blockDate = DateTime.MinValue;
            var cellForBlockDate = worksheet.Cell(blockStartRow + 1, 1);
            if (DateTime.TryParse(cellForBlockDate.GetString(), out DateTime parsedDate))
            {
                blockDate = parsedDate;
            }

            using var templateWb = new XLWorkbook(templatePath);
            var templateWs = templateWb.Worksheet(1);

            int currentTargetRow = 3;
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

            templateWs.Columns().AdjustToContents();
            foreach (var column in templateWs.ColumnsUsed()) column.Width += 8;
            templateWs.Style.Font.FontName = "Calibri";
            templateWs.Style.Font.FontSize = 10;

            int templateLastRow = templateWs.LastRowUsed()?.RowNumber() ?? 0;
            if (templateLastRow >= currentTargetRow)
            {
                for (int rowToDelete = templateLastRow; rowToDelete >= currentTargetRow; rowToDelete--)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    templateWs.Row(rowToDelete).Delete();
                }
            }

            await AddSupplementarySheetsAsync(worksheet.Workbook, templateWb, partnerName,
                new List<string> { "Activité nette à J", "J+1", "Regul", "Distributions" }, cancellationToken);

            string safePartnerName = string.Concat(partnerName.Split(Path.GetInvalidFileNameChars()));
            string dateRange = (dateStrmin == dateStrmax) ? dateStrmin : $"{dateStrmin} au {dateStrmax}";
            string outputFileName = $"COMPTE SUPPORT {safePartnerName} du {dateRange}.xlsx";
            string outputPath = Path.Combine(outputDir, outputFileName);

            templateWb.SaveAs(outputPath);
            await LogAndSend($"✅ Fichier généré pour '{partnerName}' : {outputFileName}", cancellationToken);
        }
        catch (OperationCanceledException)
        {
            await LogAndSend("❌ Génération annulée par l'utilisateur.", CancellationToken.None);
            throw;
        }
        catch (Exception ex)
        {
            string blockTitle = (blockStartRow > 0 && blockStartRow <= lastRow)
                ? worksheet.Row(blockStartRow).Cell(1).GetString()
                : "Inconnu";
            await LogAndSend($"❌ Erreur pour le bloc '{blockTitle}' (lignes {blockStartRow}-{blockEndRow}) : {ex.Message}", CancellationToken.None);
            LogOnly($"(Erreur : {ex.StackTrace})");
        }

        int currentProcessed = i - startIndex + 1;
        double percentage = (double)currentProcessed / count * 100;

        await _hubContext.Clients.All.SendAsync("ReceiveProgress", new
        {
            Current = currentProcessed,
            Total = count,
            Percentage = (int)percentage,
            Message = $"📊 {currentProcessed}/{count} fichiers générés ({(int)percentage}%)."
        }, cancellationToken);
    }

    await LogAndSend("🏁 Tous les fichiers partenaires ont été générés avec succès. Fin du processus.", cancellationToken);
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