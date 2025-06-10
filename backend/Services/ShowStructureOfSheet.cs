
using ClosedXML.Excel;

public static class ExcelUtils
{
    public static void AfficherStructureColonnes(XLWorkbook workbook, string sheetName)
    {
        var worksheet = workbook.Worksheet(sheetName);
        if (worksheet == null)
        {
            Console.WriteLine($"❌ Feuille '{sheetName}' introuvable.");
            return;
        }

        var firstRow = worksheet.Row(1);
        var secondRow = worksheet.Row(2);
        int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

        Console.WriteLine($"\n📊 Structure des colonnes (feuille '{sheetName}') :\n");

        // Vérifier si une cellule de la ligne 2 contient une date
        bool ligne2ContientDate = false;
        for (int col = 1; col <= lastColumn; col++)
        {
            var val = secondRow.Cell(col).Value;
            if (DateTime.TryParse(val.ToString(), out _))
            {
                ligne2ContientDate = true;
                break;
            }
        }

        string? currentMainTitle = null;
        var groupedColumns = new Dictionary<string, List<string>>();

        for (int col = 1; col <= lastColumn; col++)
        {
            string main = firstRow.Cell(col).GetString().Trim();

            if (!string.IsNullOrEmpty(main))
            {
                currentMainTitle = main;
                if (!groupedColumns.ContainsKey(currentMainTitle))
                    groupedColumns[currentMainTitle] = new List<string>();
            }

            if (!ligne2ContientDate)
            {
                string sub = secondRow.Cell(col).GetString().Trim();

                if (!string.IsNullOrEmpty(sub))
                {
                    if (string.IsNullOrEmpty(currentMainTitle))
                        currentMainTitle = "(Sans titre)";

                    if (!groupedColumns.ContainsKey(currentMainTitle))
                        groupedColumns[currentMainTitle] = new List<string>();

                    groupedColumns[currentMainTitle].Add(sub);
                }
                else if (!string.IsNullOrEmpty(main))
                {
                    if (!string.IsNullOrEmpty(currentMainTitle))
                    {
                        if (!groupedColumns[currentMainTitle].Contains("(Aucune sous-colonne)"))
                            groupedColumns[currentMainTitle].Add("(Aucune sous-colonne)");
                    }
                }
            }
            else
            {
                // On ignore la ligne 2 => On considère juste les titres simples (ligne 1)
                if (!string.IsNullOrEmpty(main))
                {
                    if (!string.IsNullOrEmpty(currentMainTitle))
                    {
                        if (!groupedColumns[currentMainTitle].Contains("(Aucune sous-colonne)"))
                            groupedColumns[currentMainTitle].Add("(Aucune sous-colonne)");
                    }
                }
            }
        }

        foreach (var entry in groupedColumns)
        {
            Console.WriteLine($"📁 {entry.Key}");
            foreach (var sub in entry.Value)
            {
                Console.WriteLine($"   └── 📄 {sub}");
            }
        }
    }
}