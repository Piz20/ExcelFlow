// Fichier : Models/EmailData.cs
namespace ExcelFlow.Models; // Or ExcelFlow.Utilities if you prefer

public class EmailData
{
    public string PartnerNameInFile { get; set; } = string.Empty;
    public string DateString { get; set; } = string.Empty;
    public string FinalBalance { get; set; } = string.Empty;
    public string Currency { get; set; } = string.Empty;
}