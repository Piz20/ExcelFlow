// ExcelFlow.Models/SendEmailRequest.cs
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace ExcelFlow.Models
{
    public class SendEmailRequest
    {
        // Changed from single ToEmail to separate lists for flexibility
        public List<string>? ToRecipients { get; set; }
        public List<string>? CcRecipients { get; set; }
        public List<string>? BccRecipients { get; set; }

        [Required(ErrorMessage = "Subject is required.")]
        public string Subject { get; set; } = string.Empty;

        [Required(ErrorMessage = "Body content is required.")]
        public string Body { get; set; } = string.Empty;

        public string? FromDisplayName { get; set; }

        public List<string>? AttachmentFileNames { get; set; } // List of full file paths for attachments

        // Custom validation to ensure at least one recipient is provided in any field
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