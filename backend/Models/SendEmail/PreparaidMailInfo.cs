
// Fichier : Models/SentEmailSummary.cs
using System.Collections.Generic;

namespace ExcelFlow.Models;




public class PreparedEmailInfo
{
    public string FilePath { get; set; } = string.Empty;
    public string PartnerName { get; set; } = string.Empty;
    public List<string> Emails { get; set; } = new();
    public bool IsReadyToSend { get; set; }
    public EmailData EmailData { get; set; } = new();
}
