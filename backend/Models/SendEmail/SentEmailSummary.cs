// Fichier : Models/SentEmailSummary.cs
using System.Collections.Generic;

namespace ExcelFlow.Models;

public class SentEmailSummary
{
    public string FileName { get; set; } = string.Empty;
    public string PartnerName { get; set; } = string.Empty;
    public List<string> RecipientEmails { get; set; } = new List<string>();
    public List<string> CcRecipientsSent { get; set; } = new List<string>();
    public List<string> BccRecipientsSent { get; set; } = new List<string>();
}