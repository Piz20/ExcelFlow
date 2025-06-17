// Fichier : Models/PartnerInfo.cs
using System.Collections.Generic;

namespace ExcelFlow.Models;

public class PartnerInfo
{
    public required string PartnerName { get; set; }
    public required List<string> Emails { get; set; }
}
