using System.ComponentModel.DataAnnotations;

namespace ExcelFlow.Models;

public class PartnerFileRequest
{private string _sheetName = "Analyse";

    public required string FilePath { get; set; }
    public required string TemplatePath { get; set; }
    public required string OutputDir { get; set; }

public string SheetName
{
    get => _sheetName;
    set => _sheetName = value?.Trim() ?? "Analyse";
}
    public int StartIndex { get; set; } = 0;
    public int Count { get; set; } = 3;
}
