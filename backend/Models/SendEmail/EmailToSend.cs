
namespace ExcelFlow.Models;

public class EmailToSend
{
    public string Subject { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public List<string> ToRecipients { get; set; } = new();
    public List<string> CcRecipients { get; set; } = new();
    public List<string> BccRecipients { get; set; } = new();
    public string FromDisplayName { get; set; } = string.Empty;
    public List<string> AttachmentFilePaths { get; set; } = new();
    public string? SmtpHost { get; set; }
    public int? SmtpPort { get; set; }
    public string? SmtpFromEmail { get; set; }
    public string PartnerName { get; set; } = string.Empty; // 👈 Ajout ici
}


public class PrepareEmailRequest
{
    public required string PartnerExcelPath { get; set; }
    public required string GeneratedFilesFolder { get; set; }
    public string? SmtpFromEmail { get; set; }
    public string? SmtpHost { get; set; }
    public int? SmtpPort { get; set; }
    public string? FromDisplayName { get; set; }
    public List<string>? CcRecipients { get; set; }
    public List<string>? BccRecipients { get; set; }
}


public class EmailSendResult
{
    public string To { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
}

public interface IPartnerEmailSender
{
    Task<List<EmailToSend>> PrepareCompleteEmailsAsync(
        string partnerEmailFilePath,
        string generatedFilesFolderPath,
        string? smtpFromEmail,
        string? smtpHost,
        int? smtpPort,
        string? fromDisplayName,
        List<string>? ccRecipients,
        List<string>? bccRecipients,
        CancellationToken cancellationToken);

    Task<List<EmailSendResult>> SendPreparedEmailsAsync(List<EmailToSend> emails, CancellationToken cancellationToken);


    Task SendEmailsToPartnersWithAttachments(
            string partnerEmailFilePath,
            string generatedFilesFolderPath,
            string smtpFromEmail,
            string smtpHost,
            int smtpPort,
            string fromDisplayName,
            List<string> ccRecipients,
            List<string> bccRecipients,
            CancellationToken cancellationToken);
}


