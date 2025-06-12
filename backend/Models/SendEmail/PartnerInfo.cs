// Fichier : Models/PartnerInfo.cs
using System.Collections.Generic;

namespace ExcelFlow.Models;

public class PartnerInfo
{
  // Other properties...
public required string PartnerName { get; set; }
public required List<string> Emails { get; set; }
public required string SearchableNameFull { get; set; }
public string? SearchableNameSigle { get; set; }
public required List<string> SearchableKeywords { get; set; }
}