using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelFlow.Models
{
    public class SendEmailRequest : IValidatableObject
    {
        // Listes de destinataires
        public List<string>? ToRecipients { get; set; }
        public List<string>? CcRecipients { get; set; }
        public List<string>? BccRecipients { get; set; }

        [Required(ErrorMessage = "Subject is required.")]
        public string Subject { get; set; } = string.Empty;

        [Required(ErrorMessage = "Body content is required.")]
        public string Body { get; set; } = string.Empty;

        public string? FromDisplayName { get; set; }

        public List<string>? AttachmentFileNames { get; set; } // Fichiers attachés (chemins complets)

        // --- Paramètres SMTP optionnels ---
        public string? SmtpHost { get; set; }
        public int? SmtpPort { get; set; }
        public string? SmtpFromEmail { get; set; }

        // Validation pour s'assurer qu'au moins un destinataire est présent
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if ((ToRecipients == null || !ToRecipients.Any()) &&
                (CcRecipients == null || !CcRecipients.Any()) &&
                (BccRecipients == null || !BccRecipients.Any()))
            {
                yield return new ValidationResult(
                    "At least one recipient email address must be provided in 'ToRecipients', 'CcRecipients', or 'BccRecipients'.",
                    new[] { nameof(ToRecipients), nameof(CcRecipients), nameof(BccRecipients) });
            }
        }
    }
}
