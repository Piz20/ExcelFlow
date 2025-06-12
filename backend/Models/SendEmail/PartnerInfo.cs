// Fichier : Models/PartnerInfo.cs
using System.Collections.Generic;

namespace ExcelFlow.Models;

public class PartnerInfo
{
    public string PartnerName { get; set; } = string.Empty;
    public List<string> Emails { get; set; } = new List<string>();
    public string SearchableNameFull { get; set; } = string.Empty;
    public string? SearchableNameSigle { get; set; } = null;
}